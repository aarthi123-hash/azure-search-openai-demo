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

// Enhanced header patterns
const headerPatterns = [
    /^(chapter|section|part|title|heading|overview|introduction|conclusion|summary|background|objective|scope|requirements|specifications|deliverables|timeline|budget|proposal|executive summary|technical approach|methodology|team|experience|references|appendix)\s*[\d\w\s]*[:\-]?\s*$/i,
    /^[\d]+\.?\d*\.?\s+[A-Z][^.!?]*$/,
    /^[A-Z][A-Z\s]{2,}$/,
    /^[A-Z][^.!?]*[A-Z][^.!?]*$/
];

// Title patterns for RFP/RFI documents (first page titles)
const titlePatterns = [
    /^.*REQUEST FOR (INFORMATION|PROPOSAL|QUOTATION).*$/i,
    /^.*RFP.*$/i,
    /^.*RFI.*$/i,
    /^.*RFQ.*$/i,
    /^.*GENERAL SERVICES ADMINISTRATION.*$/i,
    /^.*GSA.*$/i,
    /^.*SOLICITATION.*$/i,
    /^.*CONTRACT.*HOLDERS?.*$/i,
    /^.*ACQUISITION.*$/i,
    /^.*REGARDING.*POTENTIAL.*$/i,
    /^.*IN SUPPORT OF.*$/i,
    /^.*OFFICE OF.*$/i,
    /^.*DEPARTMENT OF.*$/i,
    /^.*AGENCY.*$/i,
    /^[A-Z][A-Z\s\(\)]{10,}$/,  // Long all-caps phrases
    /^\*\*.*\*\*$/,  // Bold markdown format
    /^.*\b(OASIS|SB|ODNI|DOD|DHS|HHS|VA|USDA)\b.*$/i,  // Common agency acronyms
];

interface HeaderParagraph {
    paragraphNumber: number;
    content: string;
}

interface HeaderSection {
    title: string;
    content: string;
    fullSection: string; // Combined title + content
    paragraphs: HeaderParagraph[]; // New: individual paragraphs with numbers
}

interface DocumentStructure {
    titles: string[];
    headerSections: HeaderSection[];
    questions: string[];
    content: string;
}

function extractTitles(text: string): string[] {
    const lines = text.split('\n');
    const titles: string[] = [];
    const firstPageLines = lines.slice(0, Math.min(50, lines.length)); // Focus on first ~50 lines
    
    for (let i = 0; i < firstPageLines.length; i++) {
        const line = firstPageLines[i].trim();
        
        // Skip empty lines or very short lines
        if (line.length < 5) continue;
        
        // Check against title patterns
        const isTitle = titlePatterns.some(pattern => pattern.test(line));
        
        // Additional checks for title-like formatting
        const isAllCaps = line === line.toUpperCase() && line.length > 10;
        const hasBoldMarkers = line.includes('**') || line.includes('*');
        const isLongCapsPhrase = /^[A-Z][A-Z\s\(\),]{15,}$/.test(line);
        const hasKeywords = /\b(REQUEST|RFP|RFI|RFQ|SOLICITATION|ACQUISITION|CONTRACT|AGENCY|OFFICE|DEPARTMENT|SUPPORT|REGARDING)\b/i.test(line);
        
        if (isTitle || isAllCaps || hasBoldMarkers || isLongCapsPhrase || (hasKeywords && line.length > 15)) {
            // Clean up the title (remove markdown formatting, extra spaces, etc.)
            let cleanTitle = line
                .replace(/^\*\*|\*\*$/g, '') // Remove markdown bold
                .replace(/\s+/g, ' ') // Normalize spaces
                .trim();
            
            // Avoid duplicates and very similar titles
            const isDuplicate = titles.some(existing => 
                existing.toLowerCase() === cleanTitle.toLowerCase() ||
                existing.toLowerCase().includes(cleanTitle.toLowerCase()) ||
                cleanTitle.toLowerCase().includes(existing.toLowerCase())
            );
            
            if (!isDuplicate && cleanTitle.length > 10) {
                titles.push(cleanTitle);
            }
        }
    }
    
    return titles.slice(0, 15); // Limit to prevent overload
}

function splitIntoParagraphs(content: string): HeaderParagraph[] {
    // Split by double line breaks or single line breaks followed by significant whitespace
    const rawParagraphs = content
        .split(/\n\s*\n|\n(?=\s{2,})/) // Split on double newlines or newline + 2+ spaces
        .map(p => p.trim())
        .filter(p => p.length > 10); // Only keep substantial paragraphs
    
    return rawParagraphs.map((paragraph, index) => ({
        paragraphNumber: index + 1,
        content: paragraph.replace(/\n/g, ' ').trim() // Clean up internal line breaks
    }));
}

function extractHeadersWithContent(text: string): HeaderSection[] {
    const lines = text.split('\n');
    const headerSections: HeaderSection[] = [];
    let currentHeader: string | null = null;
    let currentContent: string[] = [];
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        
        // Skip empty lines
        if (line.length === 0) {
            if (currentContent.length > 0) {
                currentContent.push(''); // Preserve paragraph breaks
            }
            continue;
        }
        
        // Check if this line is a header
        const isHeader = headerPatterns.some(pattern => pattern.test(line)) && line.length >= 3;
        
        if (isHeader) {
            // Save previous section if exists
            if (currentHeader !== null) {
                const contentText = currentContent.join('\n').trim();
                const paragraphs = splitIntoParagraphs(contentText);
                
                headerSections.push({
                    title: currentHeader,
                    content: contentText,
                    fullSection: `${currentHeader}\n${contentText}`,
                    paragraphs: paragraphs
                });
            }
            
            // Start new section
            currentHeader = line;
            currentContent = [];
        } else if (currentHeader !== null) {
            // Add content to current section
            currentContent.push(line);
        }
    }
    
    // Don't forget the last section
    if (currentHeader !== null) {
        const contentText = currentContent.join('\n').trim();
        const paragraphs = splitIntoParagraphs(contentText);
        
        headerSections.push({
            title: currentHeader,
            content: contentText,
            fullSection: `${currentHeader}\n${contentText}`,
            paragraphs: paragraphs
        });
    }
    
    // Filter out sections with no meaningful content
    return headerSections.filter(section => 
        section.content.length > 20 && // Must have substantial content
        !section.content.match(/^[\s\n]*$/) // Not just whitespace
    );
}

function extractQuestions(text: string) {
    const lines = text.split('\n').filter(line => line.trim().length > 0);
    const questions: string[] = [];
    
    lines.forEach(line => {
        const trimmedLine = line.trim();
        if (trimmedLine.length < 10) return;
        
        // Don't extract questions that are already part of headers or titles
        const isPartOfHeader = headerPatterns.some(pattern => pattern.test(trimmedLine));
        const isPartOfTitle = titlePatterns.some(pattern => pattern.test(trimmedLine));
        if (isPartOfHeader || isPartOfTitle) return;
        
        if (questionPatterns.some(pattern => pattern.test(trimmedLine))) {
            questions.push(trimmedLine);
        }
    });
    
    return [...new Set(questions)].slice(0, 20);
}

// Enhanced extraction using mammoth's HTML conversion for better structure detection
async function extractWithStructure(arrayBuffer: ArrayBuffer): Promise<DocumentStructure> {
    try {
        // Extract raw text for our custom parsing
        const textResult = await mammoth.extractRawText({ arrayBuffer });
        const text = textResult.value;
        
        // Extract titles from first page
        const titles = extractTitles(text);
        
        // Extract header sections with content and paragraphs
        const headerSections = extractHeadersWithContent(text);
        
        // Extract questions (excluding content already in headers and titles)
        const questions = extractQuestions(text);
        
        // Try HTML extraction for additional header detection
        try {
            const htmlResult = await mammoth.convertToHtml({ arrayBuffer });
            const htmlContent = htmlResult.value;
            
            // Extract headers from HTML structure
            const headerRegex = /<h[1-6][^>]*>(.*?)<\/h[1-6]>/gi;
            const htmlHeaders: string[] = [];
            let match;
            
            while ((match = headerRegex.exec(htmlContent)) !== null) {
                const headerText = match[1].replace(/<[^>]*>/g, '').trim();
                if (headerText.length > 0) {
                    htmlHeaders.push(headerText);
                }
            }
            
            // Add HTML headers that weren't caught by text parsing
            for (const htmlHeader of htmlHeaders) {
                const alreadyExists = headerSections.some(section => 
                    section.title.toLowerCase().includes(htmlHeader.toLowerCase()) ||
                    htmlHeader.toLowerCase().includes(section.title.toLowerCase())
                );
                
                if (!alreadyExists) {
                    // Find content after this header in the original text
                    const headerIndex = text.toLowerCase().indexOf(htmlHeader.toLowerCase());
                    if (headerIndex !== -1) {
                        const afterHeader = text.substring(headerIndex + htmlHeader.length);
                        const nextHeaderMatch = headerPatterns.find(p => {
                            const match = afterHeader.match(p);
                            return match && match.index !== undefined && match.index < 500;
                        });
                        
                        let content = '';
                        if (nextHeaderMatch) {
                            const nextHeaderIndex = afterHeader.search(nextHeaderMatch);
                            content = afterHeader.substring(0, nextHeaderIndex).trim();
                        } else {
                            content = afterHeader.substring(0, 500).trim(); // Limit to reasonable length
                        }
                        
                        if (content.length > 20) {
                            const paragraphs = splitIntoParagraphs(content);
                            headerSections.push({
                                title: htmlHeader,
                                content: content,
                                fullSection: `${htmlHeader}\n${content}`,
                                paragraphs: paragraphs
                            });
                        }
                    }
                }
            }
        } catch (htmlError) {
            console.warn("HTML extraction failed, using text-only extraction:", htmlError);
        }
        
        return {
            titles,
            headerSections: headerSections.slice(0, 25), // Limit to prevent overload
            questions,
            content: text
        };
        
    } catch (error) {
        console.error("Document extraction error:", error);
        throw error;
    }
}

export const RfpProposal: React.FC = () => {
    const [processing, setProcessing] = useState(false);
    const [results, setResults] = useState<DocumentStructure | null>(null);
    const [error, setError] = useState<string | null>(null);
    const [chatbotAnswer, setChatbotAnswer] = useState<string | null>(null);
    const [sending, setSending] = useState(false);
    const [selectedTitles, setSelectedTitles] = useState<number[]>([]);
    const [selectedQuestions, setSelectedQuestions] = useState<number[]>([]);
    const [selectedHeaderSections, setSelectedHeaderSections] = useState<number[]>([]);
    const fileInputRef = useRef<HTMLInputElement>(null);

    const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;
        
        setProcessing(true);
        setError(null);
        setResults(null);
        
        try {
            const arrayBuffer = await file.arrayBuffer();
            const extracted = await extractWithStructure(arrayBuffer);
            setResults(extracted);
            
            // Auto-select all items initially
            setSelectedTitles(extracted.titles.map((_, i) => i));
            setSelectedQuestions(extracted.questions.map((_, i) => i));
            setSelectedHeaderSections(extracted.headerSections.map((_, i) => i));

            // Send to chatbot immediately after extraction
            await sendToChatbot(createChatbotMessage(extracted));
        } catch (err) {
            setError("Error processing file. Please make sure it's a valid Word document.");
            console.error("File processing error:", err);
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
        
        // Include selected titles
        const selectedTs = selectedTitles.map(i => data.titles[i]);
        if (selectedTs.length > 0) {
            message += "DOCUMENT TITLES:\n";
            selectedTs.forEach((title, i) => {
                message += `${i + 1}. ${title}\n`;
            });
            message += "\n";
        }
        
        // Include selected header sections with content and paragraph numbers
        const selectedHS = selectedHeaderSections.map(i => data.headerSections[i]);
        if (selectedHS.length > 0) {
            message += "DOCUMENT SECTIONS WITH CONTENT:\n";
            selectedHS.forEach((section, i) => {
                message += `${i + 1}. ${section.title}\n`;
                
                // Add paragraph-by-paragraph breakdown
                if (section.paragraphs.length > 1) {
                    section.paragraphs.forEach(paragraph => {
                        message += `   Paragraph ${paragraph.paragraphNumber}: ${paragraph.content}\n`;
                    });
                } else {
                    message += `   ${section.content}\n`;
                }
                message += "\n";
            });
        }
        
        // Include selected questions
        const selectedQs = selectedQuestions.map(i => data.questions[i]);
        if (selectedQs.length > 0) {
            message += "QUESTIONS:\n";
            selectedQs.forEach((q, i) => {
                message += `${i + 1}. ${q}\n`;
            });
            message += "\n";
        }
        
        message += "Please analyze these document titles, sections (with paragraph numbers for reference), and questions and provide insights or answers. You can reference specific paragraphs using the format 'Section [number] - Paragraph [number]'.";
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
            console.error("Chatbot error:", err);
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

    // Title selection functions
    const toggleTitle = (idx: number) => {
        setSelectedTitles(selected =>
            selected.includes(idx)
                ? selected.filter(i => i !== idx)
                : [...selected, idx]
        );
    };
    
    const removeTitle = (idx: number) => {
        setSelectedTitles(selected => selected.filter(i => i !== idx));
    };

    // Header section selection functions
    const toggleHeaderSection = (idx: number) => {
        setSelectedHeaderSections(selected =>
            selected.includes(idx)
                ? selected.filter(i => i !== idx)
                : [...selected, idx]
        );
    };
    
    const removeHeaderSection = (idx: number) => {
        setSelectedHeaderSections(selected => selected.filter(i => i !== idx));
    };

    // Question selection functions
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

    return (
        <div className={styles.container}>
            <div className={styles.header}>
                <h1>📄 RFP Processor</h1>
                <p>Extract titles, complete sections with paragraph numbering, and questions from Word documents</p>
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
                        Upload a .docx or .doc file to extract titles and complete sections with numbered paragraphs
                    </p>
                </div>
                
                {processing && (
                    <div className={styles.processing}>
                        <div className={styles.spinner}></div>
                        <p>Processing document and extracting content...</p>
                    </div>
                )}
                
                {error && <div style={{ color: "red", margin: 20 }}>{error}</div>}
                
                {results && (
                    <div className={styles.results}>
                        {/* Document Titles Section */}
                        <div className={styles.section}>
                            <h3><span className={styles.sectionIcon}>🏷️</span>Document Titles</h3>
                            <div>
                                {results.titles.length > 0
                                    ? results.titles.map((title, i) => (
                                        <div className={styles.item} key={i} style={{ display: "flex", alignItems: "center", marginBottom: "10px" }}>
                                            <input
                                                type="checkbox"
                                                checked={selectedTitles.includes(i)}
                                                onChange={() => toggleTitle(i)}
                                                style={{ marginRight: 8 }}
                                            />
                                            <span style={{ flex: 1, fontWeight: "500", color: "#1e40af" }}>{title}</span>
                                            <button
                                                onClick={() => removeTitle(i)}
                                                style={{ marginLeft: 8, color: "red", background: "none", border: "none", cursor: "pointer" }}
                                                title="Remove"
                                            >
                                                ✕
                                            </button>
                                        </div>
                                    ))
                                    : <div className={styles.item}>No document titles found.</div>
                                }
                            </div>
                        </div>

                        {/* Header Sections with Content and Paragraphs */}
                        <div className={styles.section}>
                            <h3><span className={styles.sectionIcon}>📋</span>Document Sections with Paragraphs</h3>
                            <div>
                                {results.headerSections.length > 0
                                    ? results.headerSections.map((section, i) => (
                                        <div className={styles.item} key={i} style={{ display: "flex", alignItems: "flex-start", marginBottom: "20px" }}>
                                            <input
                                                type="checkbox"
                                                checked={selectedHeaderSections.includes(i)}
                                                onChange={() => toggleHeaderSection(i)}
                                                style={{ marginRight: 8, marginTop: 2 }}
                                            />
                                            <div style={{ flex: 1 }}>
                                                <div style={{ fontWeight: "600", marginBottom: "8px", color: "#2563eb" }}>
                                                    {section.title}
                                                </div>
                                                
                                                {/* Show paragraphs individually if multiple, otherwise show as single content */}
                                                {section.paragraphs.length > 1 ? (
                                                    <div style={{ marginTop: "10px" }}>
                                                        {section.paragraphs.map((paragraph, pIndex) => (
                                                            <div key={pIndex} style={{ 
                                                                marginBottom: "12px",
                                                                padding: "8px 12px",
                                                                backgroundColor: "rgba(0,0,0,0.03)",
                                                                borderRadius: "6px",
                                                                border: "1px solid rgba(0,0,0,0.08)"
                                                            }}>
                                                                <div style={{ 
                                                                    fontSize: "0.85em", 
                                                                    fontWeight: "600", 
                                                                    color: "#059669",
                                                                    marginBottom: "4px" 
                                                                }}>
                                                                    Paragraph {paragraph.paragraphNumber}:
                                                                </div>
                                                                <div style={{ 
                                                                    fontSize: "0.9em", 
                                                                    color: "#4b5563",
                                                                    lineHeight: "1.4"
                                                                }}>
                                                                    {paragraph.content}
                                                                </div>
                                                            </div>
                                                        ))}
                                                    </div>
                                                ) : (
                                                    <div style={{ 
                                                        fontSize: "0.9em", 
                                                        color: "#4b5563",
                                                        whiteSpace: "pre-wrap",
                                                        lineHeight: "1.4",
                                                        marginTop: "8px",
                                                        padding: "10px",
                                                        backgroundColor: "rgba(0,0,0,0.05)",
                                                        borderRadius: "6px",
                                                        border: "1px solid rgba(0,0,0,0.1)"
                                                    }}>
                                                        {section.content}
                                                    </div>
                                                )}
                                            </div>
                                            <button
                                                onClick={() => removeHeaderSection(i)}
                                                style={{ marginLeft: 8, color: "red", background: "none", border: "none", cursor: "pointer" }}
                                                title="Remove"
                                            >
                                                ✕
                                            </button>
                                        </div>
                                    ))
                                    : <div className={styles.item}>No sections with content found in the document.</div>
                                }
                            </div>
                        </div>

                        {/* Questions Section */}
                        <div className={styles.section}>
                            <h3><span className={styles.sectionIcon}>❓</span>Questions Found</h3>
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

                        {/* Formatted Output Section */}
                        <div className={styles.chatbotSection}>
                            <h3>📋 Extracted Content</h3>
                            <p style={{ marginBottom: 20, opacity: 0.9 }}>
                                Copy the formatted content below to use in your chatbot. Sections with multiple paragraphs include paragraph numbers for precise referencing:
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
                                            height: 300,
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

                        {/* Chatbot Interaction Section */}
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