import React, { useRef, useState } from "react";
import mammoth from "mammoth";
import { chatApi } from "../../api/api"; // Adjust path as needed
import styles from "./RfpProposal.module.css"; // Create this CSS file for custom styles
import jsPDF from "jspdf";
import { Document, Packer, Paragraph, TextRun } from "docx";
import { saveAs } from "file-saver";

const questionPatterns = [
    /^.*\?$/,
    /^(what|who|where|when|why|how|which|is|are|do|does|did|can|could|would|will|should)\s+.*$/i,
    /^(question|q[\d]*)[:\s]/i
];
const topicPatterns = [
    /^(topic|subject|chapter|section)[:\s]/i,
    /^[A-Z][^.!?]*[^.!?]$/,
    /^(discussion|analyze|explain|describe|compare)[:\s]/i
];

function extractQuestionsAndTopics(text: string) {
    const lines = text.split('\n').filter(line => line.trim().length > 0);
    const questions: string[] = [];
    const topics: string[] = [];
    lines.forEach(line => {
        const trimmedLine = line.trim();
        if (trimmedLine.length < 10) return;
        if (questionPatterns.some(pattern => pattern.test(trimmedLine))) {
            questions.push(trimmedLine);
            return;
        }
        if (topicPatterns.some(pattern => pattern.test(trimmedLine))) {
            topics.push(trimmedLine);
        }
    });
    return {
        questions: [...new Set(questions)].slice(0, 20),
        topics: [...new Set(topics)].slice(0, 15)
    };
}

export const RfpProposal: React.FC = () => {
    const [processing, setProcessing] = useState(false);
    const [results, setResults] = useState<{questions: string[], topics: string[]} | null>(null);
    const [error, setError] = useState<string | null>(null);
    const [chatbotAnswer, setChatbotAnswer] = useState<string | null>(null);
    const [sending, setSending] = useState(false);
    const [selectedQuestions, setSelectedQuestions] = useState<number[]>([]);
    const [selectedTopics, setSelectedTopics] = useState<number[]>([]);
    const fileInputRef = useRef<HTMLInputElement>(null);

    const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;
        setProcessing(true);
        setError(null);
        setResults(null);
        try {
            const arrayBuffer = await file.arrayBuffer();
            const result = await mammoth.extractRawText({ arrayBuffer });
            const text = result.value;
            const extracted = extractQuestionsAndTopics(text);
            setResults(extracted);
            setSelectedQuestions(extracted.questions.map((_, i) => i));
            setSelectedTopics(extracted.topics.map((_, i) => i));

            // Send to chatbot immediately after extraction
            await sendToChatbot(createChatbotMessage(extracted));
        } catch (err) {
            setError("Error processing file. Please make sure it's a valid Word document.");
        } finally {
            setProcessing(false);
        }
    };

    const handleUploadClick = () => {
        fileInputRef.current?.click();
    };

    const createChatbotMessage = (data = results) => {
        if (!data) return "";
        let message = "I've extracted the following from a document:\n\n";
        const selectedQs = selectedQuestions.map(i => data.questions[i]);
        const selectedTs = selectedTopics.map(i => data.topics[i]);
        if (selectedQs.length > 0) {
            message += "QUESTIONS:\n";
            selectedQs.forEach((q, i) => {
                message += `${i + 1}. ${q}\n`;
            });
            message += "\n";
        }
        if (selectedTs.length > 0) {
            message += "TOPICS:\n";
            selectedTs.forEach((t, i) => {
                message += `${i + 1}. ${t}\n`;
            });
            message += "\n";
        }
        message += "Please analyze these questions and topics and provide insights or answers.";
        return message;
    };

    const handleCopy = () => {
        navigator.clipboard.writeText(createChatbotMessage());
    };

    const sendToChatbot = async (message: string) => {
        setSending(true);
        setChatbotAnswer(null);
        try {
            const response = await chatApi({
                messages: [{ content: message, role: "user" }],
                session_state: {}
            }, false, undefined);
            const data = await response.json();
            // If the answer is an object with content, use content; otherwise fallback
            const answer =
                typeof data.answer === "object" && data.answer.content
                    ? data.answer.content
                    : typeof data.answer === "string"
                    ? data.answer
                    : typeof data.message === "object" && data.message.content
                    ? data.message.content
                    : typeof data.message === "string"
                    ? data.message
                    : "No answer received.";
            setChatbotAnswer(answer);
        } catch (err) {
            setChatbotAnswer("Error contacting chatbot.");
        } finally {
            setSending(false);
        }
    };

    const handleDownloadPdf = () => {
        if (!chatbotAnswer) return;
        const doc = new jsPDF();
        doc.setFontSize(14);
        doc.text("Chatbot Answer", 10, 20);
        doc.setFontSize(12);
        doc.text(chatbotAnswer, 10, 30, { maxWidth: 180 });
        doc.save("chatbot-answer.pdf");
    };

    const handleDownloadWord = async () => {
        if (!chatbotAnswer) return;
        const doc = new Document({
            sections: [
                {
                    properties: {},
                    children: [
                        new Paragraph({
                            children: [new TextRun({ text: "Chatbot Answer", bold: true, size: 28 })],
                        }),
                        new Paragraph({
                            children: [new TextRun({ text: chatbotAnswer, size: 24 })],
                        }),
                    ],
                },
            ],
        });
        const blob = await Packer.toBlob(doc);
        saveAs(blob, "chatbot-answer.docx");
    };

    const toggleQuestion = (idx: number) => {
        setSelectedQuestions(selected =>
            selected.includes(idx)
                ? selected.filter(i => i !== idx)
                : [...selected, idx]
        );
    };
    const removeQuestion = (idx: number) => {
        setSelectedQuestions(selected => selected.filter(i => i !== idx));
    };

    const toggleTopic = (idx: number) => {
        setSelectedTopics(selected =>
            selected.includes(idx)
                ? selected.filter(i => i !== idx)
                : [...selected, idx]
        );
    };
    const removeTopic = (idx: number) => {
        setSelectedTopics(selected => selected.filter(i => i !== idx));
    };

    return (
        <div className={styles.container}>
            <div className={styles.header}>
                <h1>📄 RFP Processor</h1>
                <p>Extract questions and topics from Word documents and send to chatbots</p>
            </div>
            <div className={styles.mainContent}>
                <div className={styles.uploadSection}>
                    <input
                        type="file"
                        accept=".docx,.doc"
                        ref={fileInputRef}
                        style={{ display: "none" }}
                        onChange={handleFileChange}
                    />
                    <button className={styles.uploadBtn} onClick={handleUploadClick}>
                        📁 Select Word Document
                    </button>
                    <p style={{ marginTop: 15, color: "#6c757d" }}>
                        Upload a .docx or .doc file to extract questions and topics
                    </p>
                </div>
                {processing && (
                    <div className={styles.processing}>
                        <div className={styles.spinner}></div>
                        <p>Processing document...</p>
                    </div>
                )}
                {error && <div style={{ color: "red", margin: 20 }}>{error}</div>}
                {results && (
                    <div className={styles.results}>
                        <div className={styles.section}>
                            <h3><span className={styles.sectionIcon}>?</span>Questions Found</h3>
                            <div>
                                {results.questions.length > 0
                                    ? results.questions.map((q, i) => (
                                        <div className={styles.item} key={i} style={{ display: "flex", alignItems: "center" }}>
                                            <input
                                                type="checkbox"
                                                checked={selectedQuestions.includes(i)}
                                                onChange={() => toggleQuestion(i)}
                                                style={{ marginRight: 8 }}
                                            />
                                            <span style={{ flex: 1 }}>{q}</span>
                                            <button
                                                onClick={() => removeQuestion(i)}
                                                style={{ marginLeft: 8, color: "red", background: "none", border: "none", cursor: "pointer" }}
                                                title="Remove"
                                            >
                                                ✕
                                            </button>
                                        </div>
                                    ))
                                    : <div className={styles.item}>No questions found in the document.</div>
                                }
                            </div>
                        </div>
                        <div className={styles.section}>
                            <h3><span className={styles.sectionIcon}>📝</span>Topics Identified</h3>
                            <div>
                                {results.topics.length > 0
                                    ? results.topics.map((t, i) => (
                                        <div className={styles.item} key={i} style={{ display: "flex", alignItems: "center" }}>
                                            <input
                                                type="checkbox"
                                                checked={selectedTopics.includes(i)}
                                                onChange={() => toggleTopic(i)}
                                                style={{ marginRight: 8 }}
                                            />
                                            <span style={{ flex: 1 }}>{t}</span>
                                            <button
                                                onClick={() => removeTopic(i)}
                                                style={{ marginLeft: 8, color: "red", background: "none", border: "none", cursor: "pointer" }}
                                                title="Remove"
                                            >
                                                ✕
                                            </button>
                                        </div>
                                    ))
                                    : <div className={styles.item}>No clear topics identified in the document.</div>
                                }
                            </div>
                        </div>
                        <div className={styles.chatbotSection}>
                            <h3>📋 Extracted Content</h3>
                            <p style={{ marginBottom: 20, opacity: 0.9 }}>
                                Copy the formatted content below to use in your chatbot:
                            </p>
                            <div className={styles.formattedOutput}>
                                <div style={{ background: "rgba(255,255,255,0.1)", borderRadius: 10, padding: 20, marginBottom: 15 }}>
                                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                                        <span style={{ fontWeight: 600 }}>Formatted Message:</span>
                                        <button className={styles.copyBtn} onClick={handleCopy}>
                                            📋 Copy
                                        </button>
                                    </div>
                                    <textarea
                                        readOnly
                                        value={createChatbotMessage()}
                                        style={{
                                            width: "100%",
                                            height: 200,
                                            background: "rgba(255,255,255,0.9)",
                                            border: "none",
                                            borderRadius: 8,
                                            padding: 15,
                                            fontFamily: "'Courier New', monospace",
                                            fontSize: "0.9rem",
                                            resize: "vertical",
                                            color: "#333"
                                        }}
                                    />
                                </div>
                            </div>
                        </div>
                        <div className={styles.chatbotInteraction}>
                            <h3>🤖 Chatbot Interaction</h3>
                            <button
                                className={styles.sendToChatbotBtn}
                                onClick={() => sendToChatbot(createChatbotMessage())}
                                disabled={sending}
                            >
                                {sending ? "Sending to Chatbot..." : "Send to Chatbot"}
                            </button>
                            {sending && <div>Sending to chatbot...</div>}
                            {chatbotAnswer && (
                                <div className={styles.section}>
                                    <h3>🤖 Chatbot Answer</h3>
                                    <div>{chatbotAnswer}</div>
                                    <div style={{ marginTop: 16 }}>
                                        <button className={styles.copyBtn} onClick={handleDownloadPdf}>
                                            📄 Download as PDF
                                        </button>
                                        <button className={styles.copyBtn} onClick={handleDownloadWord} style={{ marginLeft: 8 }}>
                                            📝 Download as Word
                                        </button>
                                    </div>
                                </div>
                            )}
                        </div>
                    </div>
                )}
            </div>
        </div>
    );
};