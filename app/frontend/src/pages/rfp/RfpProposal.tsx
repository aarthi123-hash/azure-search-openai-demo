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

// Patterns to identify scope headers and section headers
const scopeHeaderPatterns = [
    /^#+\s*(.+)$/,  // Markdown-style headers
    /^([A-Z][^.!?]*):?\s*$/,  // All caps or title case headers
    /^(scope|task|section|objective|requirement|background|purpose|introduction)\b.*$/i,
    /^[0-9]+\.?\s+([A-Z].*)/,  // Numbered headers
    /^\*\*(.+)\*\*$/,  // Bold headers in markdown
];

// Patterns to identify titles
const titlePatterns = [
    /^[A-Z][A-Z\s&()+-]{10,}$/,  // All caps titles
    /^\*\*[^*]+\*\*$/,  // Bold markdown titles
    /^#+\s*[A-Z].{5,}$/,  // Markdown headers
    /^[A-Z][^.!?]{20,}$/,  // Long title case lines
    /^(REQUEST FOR|RFP|RFI|PROPOSAL|CONTRACT|AGREEMENT|TITLE:|SUBJECT:)/i,  // Common document title prefixes
];

function extractQuestionsAndScopeContent(text: string) {
    const lines = text.split('\n').map(line => line.trim()).filter(line => line.length > 0);
    const questions: string[] = [];
    const scopeContent: string[] = [];
    const titles: string[] = [];
    
    let isInScopeSection = false;
    let currentScopeContent = "";
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        
        if (line.length < 5) continue;
        
        // Check for titles first (usually appear early in document)
        if (i < 20 && titlePatterns.some(pattern => pattern.test(line))) {
            // Clean up the title
            let cleanTitle = line.replace(/^\*\*|\*\*$/g, ''); // Remove bold markdown
            cleanTitle = cleanTitle.replace(/^#+\s*/, ''); // Remove markdown headers
            cleanTitle = cleanTitle.replace(/^>\s*/, ''); // Remove quote markers
            cleanTitle = cleanTitle.trim();
            
            if (cleanTitle.length > 5 && !titles.includes(cleanTitle)) {
                titles.push(cleanTitle);
            }
            continue;
        }
        
        // Check for questions
        if (questionPatterns.some(pattern => pattern.test(line))) {
            questions.push(line);
            continue;
        }
        
        // Check if this line is a scope/section header
        const isHeader = scopeHeaderPatterns.some(pattern => pattern.test(line));
        
        if (isHeader) {
            // If we were collecting scope content, save it before starting new section
            if (isInScopeSection && currentScopeContent.trim()) {
                scopeContent.push(currentScopeContent.trim());
                currentScopeContent = "";
            }
            
            // Check if this is a scope-related header
            const isScopeHeader = /^(scope|task|section|objective|requirement)\b.*$/i.test(line) ||
                                 line.toLowerCase().includes('scope') ||
                                 line.toLowerCase().includes('task') ||
                                 line.toLowerCase().includes('requirement');
            
            isInScopeSection = isScopeHeader;
        } else if (isInScopeSection) {
            // We're in a scope section and this is content (not a header)
            // Only add substantial content lines (not just formatting or short fragments)
            if (line.length > 20 && !line.match(/^[\-\*\+]\s*$/) && !line.match(/^[0-9]+\.\s*$/)) {
                if (currentScopeContent) {
                    currentScopeContent += " " + line;
                } else {
                    currentScopeContent = line;
                }
            }
        }
    }
    
    // Don't forget the last scope content if we were collecting it
    if (isInScopeSection && currentScopeContent.trim()) {
        scopeContent.push(currentScopeContent.trim());
    }
    
    // Break down very long scope content into smaller chunks
    const processedScopeContent: string[] = [];
    scopeContent.forEach(content => {
        if (content.length > 500) {
            // Split long content at sentence boundaries
            const sentences = content.split(/[.!?]+/).filter(s => s.trim().length > 10);
            let currentChunk = "";
            
            sentences.forEach(sentence => {
                if (currentChunk.length + sentence.length > 400) {
                    if (currentChunk.trim()) {
                        processedScopeContent.push(currentChunk.trim() + ".");
                    }
                    currentChunk = sentence.trim();
                } else {
                    currentChunk += (currentChunk ? ". " : "") + sentence.trim();
                }
            });
            
            if (currentChunk.trim()) {
                processedScopeContent.push(currentChunk.trim() + ".");
            }
        } else {
            processedScopeContent.push(content);
        }
    });
    
    return {
        questions: [...new Set(questions)].slice(0, 25),
        scopeContent: [...new Set(processedScopeContent)].slice(0, 20),
        titles: [...new Set(titles)].slice(0, 10)
    };
}

export const RfpProposal: React.FC = () => {
    const [processing, setProcessing] = useState(false);
    const [results, setResults] = useState<{questions: string[], scopeContent: string[], titles: string[]} | null>(null);
    const [error, setError] = useState<string | null>(null);
    const [chatbotAnswer, setChatbotAnswer] = useState<string | null>(null);
    const [sending, setSending] = useState(false);
    const [selectedQuestions, setSelectedQuestions] = useState<number[]>([]);
    const [selectedScopeContent, setSelectedScopeContent] = useState<number[]>([]);
    const [selectedTitles, setSelectedTitles] = useState<number[]>([]);
    const [editableQuestions, setEditableQuestions] = useState<string[]>([]);
    const [editableScopeContent, setEditableScopeContent] = useState<string[]>([]);
    const [editableTitles, setEditableTitles] = useState<string[]>([]);
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
            
            let extracted;
            if (text.length > 50000) {
                // For very large documents, process in chunks
                const chunks = [];
                for (let i = 0; i < text.length; i += 25000) {
                    chunks.push(text.substring(i, i + 25000));
                }
                
                let allQuestions: string[] = [];
                let allScopeContent: string[] = [];
                
                chunks.forEach(chunk => {
                    const chunkExtracted = extractQuestionsAndScopeContent(chunk);
                    allQuestions = [...allQuestions, ...chunkExtracted.questions];
                    allScopeContent = [...allScopeContent, ...chunkExtracted.scopeContent];
                });
                
                extracted = {
                    questions: [...new Set(allQuestions)].slice(0, 25),
                    scopeContent: [...new Set(allScopeContent)].slice(0, 20),
                    titles: [] // Large docs usually have titles at the beginning
                };
                
                setResults(extracted);
                setEditableQuestions([...extracted.questions]);
                setEditableScopeContent([...extracted.scopeContent]);
                setEditableTitles([...extracted.titles]);
                setSelectedQuestions(extracted.questions.map((_, i) => i));
                setSelectedScopeContent(extracted.scopeContent.map((_, i) => i));
                setSelectedTitles(extracted.titles.map((_, i) => i));
            } else {
                extracted = extractQuestionsAndScopeContent(text);
                setResults(extracted);
                setEditableQuestions([...extracted.questions]);
                setEditableScopeContent([...extracted.scopeContent]);
                setEditableTitles([...extracted.titles]);
                setSelectedQuestions(extracted.questions.map((_, i) => i));
                setSelectedScopeContent(extracted.scopeContent.map((_, i) => i));
                setSelectedTitles(extracted.titles.map((_, i) => i));
            }

            // Send to chatbot immediately after extraction
            await sendToChatbot(createChatbotMessage(extracted));
        } catch (err) {
            setError("Error processing file. Please make sure it's a valid Word document or the file isn't too large.");
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
        const selectedTs = selectedTitles.map(i => editableTitles[i]).filter(Boolean);
        const selectedQs = selectedQuestions.map(i => editableQuestions[i]).filter(Boolean);
        const selectedSCs = selectedScopeContent.map(i => editableScopeContent[i]).filter(Boolean);

        if (selectedTs.length > 0) {
            message += "TITLES:\n";
            selectedTs.forEach((t, i) => {
                message += `${i + 1}. ${t}\n`;
            });
            message += "\n";
        }
        if (selectedQs.length > 0) {
            message += "QUESTIONS:\n";
            selectedQs.forEach((q, i) => {
                message += `${i + 1}. ${q}\n`;
            });
            message += "\n";
        }
        if (selectedSCs.length > 0) {
            message += "SCOPE CONTENT:\n";
            selectedSCs.forEach((sc, i) => {
                message += `${i + 1}. ${sc}\n\n`;
            });
        }
        message += "Please analyze these titles, questions and scope content and provide insights or answers.";
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
        setEditableQuestions(prev => prev.filter((_, i) => i !== idx));
    };
    const updateQuestion = (idx: number, newValue: string) => {
        setEditableQuestions(prev => prev.map((q, i) => i === idx ? newValue : q));
    };

    const toggleScopeContent = (idx: number) => {
        setSelectedScopeContent(selected =>
            selected.includes(idx)
                ? selected.filter(i => i !== idx)
                : [...selected, idx]
        );
    };
    const removeScopeContent = (idx: number) => {
        setSelectedScopeContent(selected => selected.filter(i => i !== idx));
        setEditableScopeContent(prev => prev.filter((_, i) => i !== idx));
    };
    const updateScopeContent = (idx: number, newValue: string) => {
        setEditableScopeContent(prev => prev.map((sc, i) => i === idx ? newValue : sc));
    };

    const toggleTitle = (idx: number) => {
        setSelectedTitles(selected =>
            selected.includes(idx)
                ? selected.filter(i => i !== idx)
                : [...selected, idx]
        );
    };
    const removeTitle = (idx: number) => {
        setSelectedTitles(selected => selected.filter(i => i !== idx));
        setEditableTitles(prev => prev.filter((_, i) => i !== idx));
    };
    const updateTitle = (idx: number, newValue: string) => {
        setEditableTitles(prev => prev.map((t, i) => i === idx ? newValue : t));
    };

    return (
        <div className={styles.container}>
            <div className={styles.header}>
                <h1>📄 RFP Processor - Titles, Questions & Scope Content</h1>
                <p>Extract titles, questions and scope content from Word documents</p>
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
                        Upload a .docx or .doc file to extract titles, questions and scope content
                    </p>
                </div>
                {processing && (
                    <div className={styles.processing}>
                        <div className={styles.spinner}></div>
                        <p>Processing document and extracting titles, questions & scope content...</p>
                    </div>
                )}
                {error && <div style={{ color: "red", margin: 20 }}>{error}</div>}
                {results && (
                    <div className={styles.results}>
                        <div className={styles.section}>
                            <h3><span className={styles.sectionIcon}>📑</span>Titles Found</h3>
                            <div>
                                {editableTitles.length > 0
                                    ? editableTitles.map((t, i) => (
                                        <div className={styles.item} key={i} style={{ display: "flex", alignItems: "flex-start", marginBottom: 10 }}>
                                            <input
                                                type="checkbox"
                                                checked={selectedTitles.includes(i)}
                                                onChange={() => toggleTitle(i)}
                                                style={{ marginRight: 8, marginTop: 8 }}
                                            />
                                            <textarea
                                                value={t}
                                                onChange={(e) => updateTitle(i, e.target.value)}
                                                style={{
                                                    flex: 1,
                                                    minHeight: 40,
                                                    padding: 8,
                                                    border: "1px solid #ddd",
                                                    borderRadius: 4,
                                                    fontFamily: "inherit",
                                                    fontSize: "0.9rem",
                                                    resize: "vertical"
                                                }}
                                            />
                                            <button
                                                onClick={() => removeTitle(i)}
                                                style={{ marginLeft: 8, color: "red", background: "none", border: "none", cursor: "pointer", fontSize: "16px" }}
                                                title="Remove"
                                            >
                                                ✕
                                            </button>
                                        </div>
                                    ))
                                    : <div className={styles.item}>No titles found in the document.</div>
                                }
                            </div>
                        </div>
                        
                        <div className={styles.section}>
                            <h3><span className={styles.sectionIcon}>❓</span>Questions Found</h3>
                            <div>
                                {editableQuestions.length > 0
                                    ? editableQuestions.map((q, i) => (
                                        <div className={styles.item} key={i} style={{ display: "flex", alignItems: "flex-start", marginBottom: 10 }}>
                                            <input
                                                type="checkbox"
                                                checked={selectedQuestions.includes(i)}
                                                onChange={() => toggleQuestion(i)}
                                                style={{ marginRight: 8, marginTop: 8 }}
                                            />
                                            <textarea
                                                value={q}
                                                onChange={(e) => updateQuestion(i, e.target.value)}
                                                style={{
                                                    flex: 1,
                                                    minHeight: 40,
                                                    padding: 8,
                                                    border: "1px solid #ddd",
                                                    borderRadius: 4,
                                                    fontFamily: "inherit",
                                                    fontSize: "0.9rem",
                                                    resize: "vertical"
                                                }}
                                            />
                                            <button
                                                onClick={() => removeQuestion(i)}
                                                style={{ marginLeft: 8, color: "red", background: "none", border: "none", cursor: "pointer", fontSize: "16px" }}
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
                            <h3><span className={styles.sectionIcon}>📋</span>Scope Content</h3>
                            <div>
                                {editableScopeContent.length > 0
                                    ? editableScopeContent.map((sc, i) => (
                                        <div className={styles.item} key={i} style={{ display: "flex", alignItems: "flex-start", marginBottom: 10 }}>
                                            <input
                                                type="checkbox"
                                                checked={selectedScopeContent.includes(i)}
                                                onChange={() => toggleScopeContent(i)}
                                                style={{ marginRight: 8, marginTop: 8 }}
                                            />
                                            <textarea
                                                value={sc}
                                                onChange={(e) => updateScopeContent(i, e.target.value)}
                                                style={{
                                                    flex: 1,
                                                    minHeight: 80,
                                                    padding: 8,
                                                    border: "1px solid #ddd",
                                                    borderRadius: 4,
                                                    fontFamily: "inherit",
                                                    fontSize: "0.9rem",
                                                    lineHeight: 1.4,
                                                    resize: "vertical"
                                                }}
                                            />
                                            <button
                                                onClick={() => removeScopeContent(i)}
                                                style={{ marginLeft: 8, color: "red", background: "none", border: "none", cursor: "pointer", fontSize: "16px" }}
                                                title="Remove"
                                            >
                                                ✕
                                            </button>
                                        </div>
                                    ))
                                    : <div className={styles.item}>No scope content found in the document.</div>
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
                                            height: 250,
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