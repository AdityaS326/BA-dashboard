// backend/controllers/transcribeController.js
// Transcribes audio/video via Groq Whisper with speaker diarization.
// Large/video files are converted to 64 kbps mono MP3 via ffmpeg, chunked if needed.
// Speaker labels are added by the Groq LLM from conversation patterns.

import { execFile }                                        from "child_process";
import { promisify }                                       from "util";
import { writeFile, readFile, rm, mkdtemp, readdir }       from "fs/promises";
import { tmpdir }                                          from "os";
import { join, extname }                                   from "path";
import ffmpegPath                                          from "ffmpeg-static";

const execFileAsync = promisify(execFile);

const WHISPER_LIMIT       = 24 * 1024 * 1024;
const GROQ_WHISPER_URL    = "https://api.groq.com/openai/v1/audio/transcriptions";
const GROQ_CHAT_URL       = "https://api.groq.com/openai/v1/chat/completions";
const GROQ_WHISPER_MODEL  = "whisper-large-v3";
const WHISPER_PROMPT      = "Teams meeting recording. Multiple speakers discuss project progress, UI, features, and tasks. Speakers may have Indian accents.";

function fmtTime(secs) {
  const m = Math.floor(secs / 60);
  const s = Math.floor(secs % 60).toString().padStart(2, "0");
  return `${m}:${s}`;
}

// Transcribe one audio buffer → { text, segments } with timestamps offset by offsetSecs
async function transcribeSegments(buffer, filename, apiKey, offsetSecs = 0) {
  const form = new FormData();
  form.append("file", new Blob([buffer], { type: "audio/mpeg" }), filename);
  form.append("model",           GROQ_WHISPER_MODEL);
  form.append("response_format", "verbose_json");
  form.append("prompt",          WHISPER_PROMPT);

  const resp = await fetch(GROQ_WHISPER_URL, {
    method:  "POST",
    headers: { Authorization: `Bearer ${apiKey}` },
    body:    form,
  });
  if (!resp.ok) throw new Error(`Transcription error: ${await resp.text()}`);
  const data = await resp.json();

  const segments = (data.segments || []).map(s => ({
    start: (s.start || 0) + offsetSecs,
    end:   (s.end   || 0) + offsetSecs,
    text:  (s.text  || "").trim(),
  }));

  return { text: data.text || "", segments };
}

// Use Groq LLM to assign speaker labels from conversation patterns
async function addSpeakerLabels(segments, apiKey) {
  if (!segments.length) return "";

  const rawLines = segments
    .filter(s => s.text)
    .map(s => `[${fmtTime(s.start)}] ${s.text}`)
    .join("\n");

  const prompt = `You are formatting a meeting transcript. Identify distinct speakers purely from conversational cues: questions vs answers, "yes/no/okay" replies, addressing someone by name, topic shifts, and turn-taking patterns.

Transcript with timestamps:
${rawLines}

Output format (strict):
Speaker 1 (M:SS):
[everything this speaker said in this turn]

Speaker 2 (M:SS):
[everything this speaker said in this turn]

Rules:
- Use "Speaker 1", "Speaker 2" etc. (do not invent names unless a name is spoken in the audio)
- Merge consecutive lines from the same speaker into one block
- Separate each speaker block with a blank line
- Return ONLY the formatted transcript — no explanations, no commentary`;

  const resp = await fetch(GROQ_CHAT_URL, {
    method:  "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey}` },
    body: JSON.stringify({
      model:       process.env.GROQ_MODEL || "llama-3.3-70b-versatile",
      messages:    [{ role: "user", content: prompt }],
      max_tokens:  4000,
      temperature: 0.1,
    }),
  });

  if (!resp.ok) return rawLines; // fallback: timestamped plain text
  const data = await resp.json();
  return data.choices?.[0]?.message?.content?.trim() || rawLines;
}

export async function transcribeAudio(req, res) {
  if (!req.file) return res.status(400).json({ error: "No file uploaded." });

  const apiKey = process.env.GROQ_API_KEY;
  if (!apiKey) return res.status(500).json({ error: "GROQ_API_KEY not set in .env" });

  const isVideo = /\.(mp4|mov|webm|mkv)$/i.test(req.file.originalname);

  // ── Small audio file: transcribe directly ─────────────────────
  if (req.file.size <= WHISPER_LIMIT && !isVideo) {
    try {
      const result     = await transcribeSegments(req.file.buffer, req.file.originalname, apiKey);
      const transcript = await addSpeakerLabels(result.segments, apiKey);
      return res.json({ transcript: transcript || result.text });
    } catch (err) {
      console.error("[transcribeController]", err.message);
      return res.status(502).json({ error: err.message });
    }
  }

  // ── Large file / video: convert → chunk → transcribe ──────────
  const tmpDir    = await mkdtemp(join(tmpdir(), "ba-whisper-"));
  const inputExt  = extname(req.file.originalname) || ".mp4";
  const inputPath = join(tmpDir, `input${inputExt}`);
  const mp3Path   = join(tmpDir, "audio.mp3");

  try {
    await writeFile(inputPath, req.file.buffer);

    // Convert to 64 kbps mono 16 kHz MP3 (good speech quality, ~10× smaller than MP4)
    await execFileAsync(ffmpegPath, [
      "-i", inputPath,
      "-vn",
      "-acodec", "libmp3lame",
      "-ab", "64k",
      "-ar", "16000",
      "-ac", "1",
      "-y", mp3Path,
    ]);

    const mp3Buffer  = await readFile(mp3Path);
    let allSegments  = [];

    if (mp3Buffer.length <= WHISPER_LIMIT) {
      const result = await transcribeSegments(mp3Buffer, "audio.mp3", apiKey, 0);
      allSegments  = result.segments;
      console.log(`[transcribeController] converted: ${req.file.size} → ${mp3Buffer.length} bytes, ${allSegments.length} segments`);
    } else {
      // Split into 20-minute chunks
      console.log(`[transcribeController] splitting into chunks (${mp3Buffer.length} bytes)`);
      await execFileAsync(ffmpegPath, [
        "-i", mp3Path,
        "-f", "segment",
        "-segment_time", "1200",
        "-c", "copy",
        "-y", join(tmpDir, "chunk_%03d.mp3"),
      ]);

      const chunkFiles = (await readdir(tmpDir))
        .filter(f => f.startsWith("chunk_") && f.endsWith(".mp3"))
        .sort();

      let offsetSecs = 0;
      for (const f of chunkFiles) {
        const buf    = await readFile(join(tmpDir, f));
        const result = await transcribeSegments(buf, f, apiKey, offsetSecs);
        allSegments.push(...result.segments);
        offsetSecs += 1200;
      }
    }

    const transcript = await addSpeakerLabels(allSegments, apiKey);
    res.json({ transcript });
  } catch (err) {
    console.error("[transcribeController] error:", err.message);
    res.status(502).json({ error: `Transcription failed: ${err.message}` });
  } finally {
    await rm(tmpDir, { recursive: true, force: true }).catch(() => {});
  }
}
