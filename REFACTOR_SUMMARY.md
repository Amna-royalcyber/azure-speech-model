# Teams Media Bot Refactor Summary

This document summarizes the end-to-end refactor completed in this repository to move from post-transcription identity handling to an identity-first, SSRC/sourceId-based pipeline.

## Goal

Refactor the bot so identity is determined before speech transcription:

- No AWS Transcribe path
- No diarization / speaker guessing / post-hoc identity backfill
- Per-stream processing (Teams unmixed sourceId stream)
- Deterministic stream-to-user mapping
- Azure Speech used as text conversion only

## Final Pipeline (Current)

1. Graph participant roster events register users immediately.
2. Graph participant `mediaStreams` are parsed to collect `sourceId` values.
3. `sourceId` is treated as SSRC-like stream id and bound to participant identity.
4. Media ingestion reads each unmixed buffer, extracts stream id at ingestion boundary.
5. Router resolves stream id to participant identity before calling Azure Speech.
6. If stream is temporarily unmapped, audio is buffered briefly (3s max) and replayed only after mapping exists; stale frames are dropped.
7. Azure Speech recognizer runs per stream id and emits structured transcript events.
8. Broadcaster forwards structured transcripts to SignalR without identity remapping.
9. Chunk manager stores/transmits transcript chunks without identity rewriting.

## Major Removals

The following legacy components were removed earlier in the refactor:

- AWS transcribe services and stream managers
- mixed/dominant delay gates
- legacy transcript aggregation and dedupe layers
- speaker identity backfill store
- hosted identity backfill worker
- post-hoc identity correction paths

Specific files removed in this repo:

- `AwsTranscribeService.cs`
- `TranscribeStreamService.cs`
- `TranscriptionManager.cs`
- `ParticipantAudioStreamHandler.cs`
- `MixedDominantAudioDelayGate.cs`
- `UnmixedAudioDelayGate.cs`
- `TranscriptAggregator.cs`
- `TranscriptBuffer.cs`
- `TranscriptDeduplicator.cs`
- `TranscriptIdentityResolver.cs`
- `IdentityBackfillService.cs`
- `SpeakerIdentityStore.cs`

## New / Added Files

- `TranscriptionParticipant.cs`
  - Strong identity object passed to transcription (`ParticipantId`, `DisplayName`, `IntraId`).

- `UnmixedAudioHelpers.cs`
  - Extracts SSRC/sourceId from `UnmixedAudioBuffer`.
  - Copies unmanaged payload into managed byte array.

- `REFACTOR_SUMMARY.md` (this document)

## Core File Changes

### `GraphParticipantMediaStreams.cs`

Added identity mapping structures:

- `MediaStreamInfo`
- `SsrcParticipantMapper`
  - `Bind(uint ssrc, string participantId)`
  - `GetParticipantIdBySsrc(uint ssrc)`
  - `GetParticipantId(uint ssrc)` alias
  - `HasMapping(uint ssrc)`
  - `GetSsrcToParticipantMap()`
  - `RemoveSsrc(uint ssrc)`
  - `Clear()`

Kept robust JSON parsing for Graph `mediaStreams` extraction.

### `MeetingParticipantService.cs`

Identity registration and binding were tightened:

- Added immediate participant registration method:
  - `AddOrUpdateParticipant(participantId, displayName, intraId)`
- Added stream bind method:
  - `BindMediaStreamToParticipant(sourceId, entraObjectId, intraCallParticipantId)`
- Added transcription participant resolver:
  - `TryGetTranscriptionParticipant(sourceId, out TranscriptionParticipant)`
- On Graph participant updates:
  - register participant immediately
  - parse and bind all `sourceId` values to identity
- No blocking registration if `sourceId` is missing:
  - logs warning and continues
- Removed late single-participant auto-binding logic (strict identity mode was re-enforced before safe buffering was added).

### `MediaHandler.cs`

Moved stream identity extraction to ingestion boundary:

- No longer forwards raw `AudioMediaReceivedEventArgs` to router
- For each unmixed buffer:
  - extracts `ssrc/sourceId` via `UnmixedAudioHelpers.TryGetSsrc`
  - copies payload
  - calls router:
    - `HandleAudioAsync(ssrc, payload, timestamp)`
- Note: stream id is read from each `UnmixedAudioBuffer` (Teams unmixed model), not from a single top-level buffer source id.

### `ParticipantAudioRouter.cs`

This file now enforces identity-first routing with safe temporary buffering:

- Receives explicit stream id and audio payload:
  - `HandleAudioAsync(uint ssrc, byte[] rawPayload, long timestampHns)`
- Periodically refreshes participant stream mappings from call roster.
- Checks mapping before processing.
- Safe temporary buffer for unmapped streams:
  - per-SSRC in-memory queue
  - timeout: 3 seconds
  - stale frames removed during enqueue/replay
  - buffered frames are replayed in order once mapping exists
- On mapping availability:
  - flushes buffered frames first
  - then processes current frame
- Converts audio to PCM and forwards to Azure Speech only when identity is available.
- Unmapped stream logging is throttled to avoid log spam.

### `AzureSpeechTranscriptionService.cs`

Refactored to identity-injected processing:

- Public API now expects identity:
  - `ProcessAudioAsync(uint ssrc, TranscriptionParticipant participant, byte[] pcm16kMono, long timestampHns)`
- One recognizer per stream id.
- Recognizer events do not perform identity inference.
- Transcript emit includes identity from `TranscriptionParticipant`.
- Added recognition telemetry log:
  - `TRANSCRIPT [DisplayName]: text`
- `BroadcastStructuredTranscriptAsync(...)` called with pre-resolved identity.
- Chunk manager called with participant identity as provided.

### `TranscriptBroadcaster.cs`

Simplified to forward-only behavior:

- Broadcasts structured transcript with:
  - `intraId`
  - `participantId`
  - `displayName`
  - `ssrc` + `sourceId` alias
  - `text`
  - `confidence`
  - `timestamp`
- Added compatibility method:
  - `BroadcastTranscriptIdentityUpdateAsync(...)` (for external stale callers)
- No participant identity remapping logic in broadcaster.

### `TranscriptionChunkManager.cs`

Removed identity rewrite dependency:

- No longer depends on `IParticipantManager` for label/identity mutation during chunk flush.
- Uses participant identity values already attached to transcript items.

### `Program.cs`

DI cleanup and registration updates:

- Removed deleted services from DI:
  - `SpeakerIdentityStore`
  - `IdentityBackfillService`
- Added:
  - `SsrcParticipantMapper`
- Join endpoint update:
  - returns HTTP 409 Conflict when attempting join while a call is already active.

### `CallHandler.cs`

- Removed `SpeakerIdentityStore` dependency and reset calls.
- Kept active-call guard (`already active`) behavior.

### `ParticipantManager.cs`

- Removed scope-factory / identity-store reflection path.
- Constructor now uses only `ILogger<ParticipantManager>`.

### `TeamsMediaBot.csproj`

- Removed AWS package references.
- Added Azure Speech package:
  - `Microsoft.CognitiveServices.Speech`

### `appsettings.json`

- Added Azure speech config keys under `Bot`:
  - `AzureSpeechKey`
  - `AzureSpeechRegion`
- Removed AWS region section.

### `wwwroot/index.html`

Updated SignalR transcript rendering compatibility:

- Supports `displayName` and `speakerLabel`
- Uses `ssrc/sourceId` metadata
- Keeps transcript display robust when labels are absent

## Behavior Notes

### Mapping and buffering

- Identity mapping is still authoritative from Graph `mediaStreams` source ids.
- Router now has a short safe pre-map buffer (3 seconds) to absorb Graph/sourceId timing delays.
- Frames older than timeout are discarded to prevent unbounded memory growth.
- Audio is never transcribed without identity: buffering only delays processing until stream mapping is available.

### No post-processing identity assignment

- There is no diarization-based speaker reassignment.
- There is no transcript-time identity guessing.
- Azure Speech does not decide speaker identity.

### Active call join guard

- If `/api/bot/join` is called while call is active, API now responds with conflict semantics instead of generic 500.

## Build Status

Local project build passes after refactor:

- `dotnet build` succeeds.

## Current Runtime Diagnostics To Watch

Helpful log patterns during live testing:

- SSRC bind from Graph:
  - `[SSRC BIND] sourceId ... -> ...`
- Buffered due to missing mapping:
  - `Buffering audio: SSRC/sourceId ... is not mapped yet...`
- Speech output confirmed:
  - `TRANSCRIPT [DisplayName]: ...`
- No unmixed stream warning:
  - `No UnmixedAudioBuffers in this frame ...`

## Deployment Reminder

If VM build/runtime path differs from this repo path, deploy/sync this full updated tree to the VM path before running. Several earlier runtime errors were caused by mixed old/new files in different folders.
