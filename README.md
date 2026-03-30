# Whisper Windows Transcriber (GUI)

Simple Windows UI app for running local whisper transcription and saving one `.txt` per audio file.

## Files

- `WhisperTranscriber.ps1`: Main Windows UI app.
- `run-transcriber.cmd`: Double-click launcher.

## What it does

- Lets you choose:
  - Whisper executable (`whisper-cli.exe` recommended)
  - Whisper model `.bin` (for example large-v3 Q5/Q8)
  - Transcription language from a supported Whisper language dropdown
  - Optional translate-to-English mode (`-tr`)
  - Audio input in two batch modes:
    - add multiple files
    - add a folder of audio files (optional recursive include)
  - Output folder (only folder selection is needed)
  - Prompt text (pre-filled for speaker-wise formatting)
- Automatically creates transcript names from audio filenames:
  - `meeting.mp3` -> `meeting.txt`
  - if duplicate basenames appear in one batch, it appends `_2`, `_3`, etc.
- Optionally enables diarization (`-tdrz`) and no timestamps (`-nt`).
- Shows supported file formats in the UI above audio source selection.
- Auto-converts `m4a/aac/wma/opus/mp4/m4b` to temporary WAV using `ffmpeg` before transcription.
- At startup, warns if `ffmpeg` is missing and offers to install with `winget`.

## Setup (First Time)

1. Download a Windows build of `whisper.cpp` that includes `whisper-cli.exe`.
   - Use the official `whisper.cpp` GitHub Releases page and pick the latest Windows binary package.
2. Extract it (example: `D:\audio-transcription-whisper-windows\whisper-bin-x64\Release\whisper-cli.exe`).
3. Download a Whisper model `.bin` file (example: `ggml-large-v3-q5_0.bin` or `ggml-large-v3-q8_0.bin`).
4. In the app, set:
   - `Whisper EXE` -> your `whisper-cli.exe`
   - `Model (.bin)` -> your downloaded model file

## ffmpeg Installation

`ffmpeg` is needed for non-native formats (`m4a/aac/wma/opus/mp4/m4b`).

You can install it either way:

1. In-app install (recommended):
   - Start the app.
   - If `ffmpeg` is missing, you will see a popup:
     - `Yes`: app installs ffmpeg automatically using `winget`
     - `No`: you install it manually later
2. Manual install:
   - Run:
     ```powershell
     winget install Gyan.FFmpeg
     ```
   - Or place `ffmpeg.exe` next to `whisper-cli.exe`
   - Or add `ffmpeg` to your system PATH

## Run

1. Double-click `run-transcriber.cmd`
2. In the app, select:
   - Whisper executable path (`whisper-cli.exe`)
   - Model file (`.bin`)
   - Audio files or folder
   - Output folder
3. Click **Start Transcription**

## Model note

Use a large-v3 quantized model when available, such as:

- `ggml-large-v3.bin`
- `ggml-large-v2-q8_0.bin` (Q8)

You can download the model later and then select it in the UI.

## Prompt used by default

`Transcribe this audio faithfully. Segment by speaker turns and format as Speaker 1:, Speaker 2:, etc. Preserve meaning, key pauses, hesitations, and sentence boundaries.`

You can edit this prompt in the UI before each run.

## Troubleshooting

- If the executable is not found, browse manually to `whisper-cli.exe`.
- If transcription fails for a compressed format, make sure `ffmpeg` is installed (or use WAV/MP3/OGG/FLAC).
- If speaker labels are weak, keep `-tdrz` enabled and refine the prompt wording.
- If you use `m4a/aac/wma/opus/mp4/m4b`, ensure `ffmpeg.exe` is available in PATH or next to `whisper-cli.exe`.
