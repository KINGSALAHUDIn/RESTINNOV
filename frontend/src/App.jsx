import { BrowserRouter, Routes, Route } from "react-router-dom";
import AudioRecorder from "./AudioRecorder";
import TranscriptResult from "./TranscriptResult";

export default function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route path="/" element={<AudioRecorder />} />
        <Route path="/result" element={<TranscriptResult />} />
      </Routes>
    </BrowserRouter>
  );
}