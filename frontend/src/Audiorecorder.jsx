import { useState, useRef, useEffect } from "react";
import { useNavigate } from "react-router-dom";
import "./AudioRecorder.css";

export default function AudioRecorder() {
  const [status, setStatus] = useState("idle"); // idle | recording | paused | transcribing
  const [error, setError] = useState("");
  const [seconds, setSeconds] = useState(0);
  const navigate = useNavigate();

  const mediaRecorderRef = useRef(null);
  const chunksRef = useRef([]);
  const streamRef = useRef(null);
  const timerRef = useRef(null);
  const analyserRef = useRef(null);
  const animFrameRef = useRef(null);
  const barsRef = useRef([]);

  const fmt = (s) =>
    String(Math.floor(s / 60)).padStart(2, "0") +
    ":" +
    String(s % 60).padStart(2, "0");

  function drawBars() {
    if (!analyserRef.current) return;
    const data = new Uint8Array(analyserRef.current.frequencyBinCount);
    analyserRef.current.getByteFrequencyData(data);
    barsRef.current.forEach((bar, i) => {
      if (!bar) return;
      const idx = Math.floor((i * data.length) / barsRef.current.length / 2);
      const v = data[idx];
      const h = Math.max(3, Math.round((v * 56) / 255));
      bar.style.height = h + "px";
      bar.style.opacity = v > 60 ? "1" : "0.4";
    });
    animFrameRef.current = requestAnimationFrame(drawBars);
  }

  function resetBars() {
    barsRef.current.forEach((bar) => {
      if (bar) {
        bar.style.height = "3px";
        bar.style.opacity = "0.3";
      }
    });
  }

  async function startRecording() {
    setError("");
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      streamRef.current = stream;

      const ctx = new (window.AudioContext || window.webkitAudioContext)();
      const src = ctx.createMediaStreamSource(stream);
      const analyser = ctx.createAnalyser();
      analyser.fftSize = 256;
      src.connect(analyser);
      analyserRef.current = analyser;

      const recorder = new MediaRecorder(stream);
      mediaRecorderRef.current = recorder;
      chunksRef.current = [];

      recorder.ondataavailable = (e) => chunksRef.current.push(e.data);
      recorder.onstop = handleStop;
      recorder.start();

      setStatus("recording");
      setSeconds(0);
      timerRef.current = setInterval(() => setSeconds((s) => s + 1), 1000);
      animFrameRef.current = requestAnimationFrame(drawBars);
    } catch {
      setError("Microphone access denied. Please allow microphone permission.");
    }
  }

  function pauseRecording() {
    if (mediaRecorderRef.current?.state === "recording") {
      mediaRecorderRef.current.pause();
      clearInterval(timerRef.current);
      cancelAnimationFrame(animFrameRef.current);
      resetBars();
      setStatus("paused");
    }
  }

  function resumeRecording() {
    if (mediaRecorderRef.current?.state === "paused") {
      mediaRecorderRef.current.resume();
      timerRef.current = setInterval(() => setSeconds((s) => s + 1), 1000);
      animFrameRef.current = requestAnimationFrame(drawBars);
      setStatus("recording");
    }
  }

  function stopRecording() {
    mediaRecorderRef.current?.stop();
    streamRef.current?.getTracks().forEach((t) => t.stop());
    clearInterval(timerRef.current);
    cancelAnimationFrame(animFrameRef.current);
    resetBars();
    setStatus("transcribing");
  }

  async function handleStop() {
    const blob = new Blob(chunksRef.current, { type: "audio/webm" });
    await sendToFlask(blob);
  }

  async function sendToFlask(blob) {
    const formData = new FormData();
    formData.append("audio", blob, "recording.webm");

    try {
      const res = await fetch("http://localhost:5000/transcribe", {
        method: "POST",
        body: formData,
      });

      if (!res.ok) throw new Error(`Server error: ${res.status}`);
      const data = await res.json();

      if (data.error) throw new Error(data.error);

      // Navigate to result page and pass all inspection data
      navigate("/result", { 
        state: { 
          transcript: data.transcript,
          inspection: data.inspection,
          saved: data.saved,
          id: data.id
        } 
      });
    } catch (err) {
      setError("Transcription failed: " + err.message);
      setStatus("idle");
    }
  }

  useEffect(() => {
    return () => {
      clearInterval(timerRef.current);
      cancelAnimationFrame(animFrameRef.current);
    };
  }, []);

  const isRecording = status === "recording";
  const isPaused = status === "paused";
  const isTranscribing = status === "transcribing";
  const isActive = isRecording || isPaused;

  return (
    <div className="ar-page">
      <div className="ar-card">
        <div className="ar-header">
          <span className="ar-label">Voice to Text</span>
          <h1 className="ar-title">Audio Recorder</h1>
        </div>

        <div className="ar-waveform">
          {Array.from({ length: 36 }).map((_, i) => (
            <div
              key={i}
              className={`ar-bar ${isRecording ? "ar-bar--active" : ""}`}
              ref={(el) => (barsRef.current[i] = el)}
            />
          ))}
        </div>

        <div className={`ar-timer ${isRecording ? "ar-timer--live" : ""}`}>
          {isTranscribing ? (
            <span className="ar-transcribing-text">Sending…</span>
          ) : (
            fmt(seconds)
          )}
        </div>

        <div className="ar-controls">
          {!isActive && !isTranscribing && (
            <button className="ar-btn ar-btn--record" onClick={startRecording}>
              <span className="ar-btn__dot" />
              {status === "idle" && seconds === 0 ? "Start Recording" : "New Recording"}
            </button>
          )}

          {isRecording && (
            <>
              <button className="ar-btn ar-btn--ghost" onClick={pauseRecording}>
                <PauseIcon /> Pause
              </button>
              <button className="ar-btn ar-btn--stop" onClick={stopRecording}>
                <StopIcon /> Send to Whisper
              </button>
            </>
          )}

          {isPaused && (
            <>
              <button className="ar-btn ar-btn--ghost" onClick={resumeRecording}>
                <PlayIcon /> Resume
              </button>
              <button className="ar-btn ar-btn--stop" onClick={stopRecording}>
                <StopIcon /> Send to Whisper
              </button>
            </>
          )}

          {isTranscribing && (
            <div className="ar-spinner-wrap">
              <div className="ar-spinner" />
            </div>
          )}
        </div>

        <div className={`ar-status ${isRecording ? "ar-status--live" : ""}`}>
          {isRecording && <span className="ar-dot-live" />}
          {isRecording
            ? "recording"
            : isPaused
            ? "paused"
            : isTranscribing
            ? "sending to Whisper…"
            : "ready"}
        </div>

        {error && <div className="ar-error">{error}</div>}
      </div>
    </div>
  );
}

function PauseIcon() {
  return (
    <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
      <rect x="2" y="1" width="3.5" height="12" rx="1" />
      <rect x="8.5" y="1" width="3.5" height="12" rx="1" />
    </svg>
  );
}
function StopIcon() {
  return (
    <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
      <rect x="2" y="2" width="10" height="10" rx="2" />
    </svg>
  );
}
function PlayIcon() {
  return (
    <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
      <polygon points="2,1 13,7 2,13" />
    </svg>
  );
}