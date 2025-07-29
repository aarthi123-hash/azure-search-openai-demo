import React, { useRef, useState } from "react";
import mammoth from "mammoth";
import { chatApi } from "../../api/api"; // Adjust path as needed
import styles from "./RfpProposal.module.css"; // Create this CSS file for custom styles
import jsPDF from "jspdf";
import { Document, Packer, Paragraph, TextRun, PageBreak, Header, Footer, AlignmentType } from "docx";
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
    /^[A-Z][^.!?]*[A-Z][^.!?]*$/,
    /^\d+\.\d+\s+.+$/, // Matches "3.1 Task Title", "3.2 Another Task", etc.
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
    const [qaStream, setQaStream] = useState<{ question: string; answer: string | null }[]>([]);
    const [currentlyProcessing, setCurrentlyProcessing] = useState<number | null>(null);
    const [newQuestion, setNewQuestion] = useState("");
    const [addingQuestion, setAddingQuestion] = useState(false);
    const fileInputRef = useRef<HTMLInputElement>(null);
    const [selectedParagraphs, setSelectedParagraphs] = useState<{ [key: string]: boolean }>({});


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
    if (!qaStream.length) return;
    
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.width;
    const pageHeight = doc.internal.pageSize.height;
    const margin = 20;
    const maxWidth = pageWidth - (margin * 2);
    
    // Header styling
    doc.setFontSize(18);
    doc.setFont("helvetica", 'bold');
    doc.text('Chatbot Conversation Report', margin, 25);
    
    // Metadata with better spacing
    doc.setFontSize(10);
    doc.setFont("helvetica", 'normal');
    doc.text(`User: Guest`, margin, 40);
    doc.text(`Date: ${new Date().toLocaleDateString()}`, margin, 48);
    doc.text(`Time: ${new Date().toLocaleTimeString()}`, margin, 56);
    doc.text(`Total Questions: ${qaStream.length}`, margin, 64);
    
    // Add a separator line
    doc.setLineWidth(0.5);
    doc.line(margin, 70, pageWidth - margin, 70);
    
    let y = 80;
    
    qaStream.forEach((qa, idx) => {
        // Check if we need a new page
        if (y > pageHeight - 60) {
            doc.addPage();
            y = 30;
        }
        
        // Question styling
        doc.setFontSize(11);
        doc.setFont("helvetica", 'normal');
        doc.setTextColor(0, 0, 0);
        
        const questionText = `Q${idx + 1}: ${qa.question}`;
        const questionLines = doc.splitTextToSize(questionText, maxWidth);
        doc.text(questionLines, margin, y);
        y += (questionLines.length * 6) + 2;
        
        // Answer styling
        doc.setFontSize(10);
        doc.setFont("helvetica", 'normal');
        doc.setTextColor(60, 60, 60);
        
        const answerText = `A${idx + 1}: ${qa.answer ?? "No answer provided"}`;
        const answerLines = doc.splitTextToSize(answerText, maxWidth);
        doc.text(answerLines, margin, y);
        y += (answerLines.length * 5) + 8;
        
        // Add a subtle separator between Q&A pairs
        if (idx < qaStream.length - 1) {
            doc.setDrawColor(200, 200, 200);
            doc.setLineWidth(0.1);
            doc.line(margin, y, pageWidth - margin, y);
            y += 8;
        }
    });
    
    // Footer on each page
    const totalPages = doc.getNumberOfPages();
    for (let i = 1; i <= totalPages; i++) {
        doc.setPage(i);
        doc.setFontSize(8);
        doc.setTextColor(128, 128, 128);
        doc.text(`Page ${i} of ${totalPages}`, pageWidth - 40, pageHeight - 10);
        doc.text(`Generated on ${new Date().toLocaleString()}`, margin, pageHeight - 10);
    }
    
    doc.save("chatbot-conversation-report.pdf");
};

const handleDownloadWord = async () => {
    // Load previous Q&A from local storage
    let previousQaStream: { question: string; answer: string | null }[] = [];
    try {
        const stored = localStorage.getItem("rfpQaStream");
        if (stored) previousQaStream = JSON.parse(stored);
    } catch {}

    // Merge previous and current, avoiding duplicates
    const allQaStream = [
        ...previousQaStream,
        ...qaStream.filter(
            qa => !previousQaStream.some(
                prev => prev.question === qa.question && prev.answer === qa.answer
            )
        )
    ];

    if (!allQaStream.length) return;

    // Save merged Q&A back to local storage
    localStorage.setItem("rfpQaStream", JSON.stringify(allQaStream));

    const doc = new Document({
        sections: [
            {
                properties: {
                    page: {
                        margin: {
                            top: 1440,    // 1 inch
                            right: 1440,  // 1 inch
                            bottom: 1440, // 1 inch
                            left: 1440,   // 1 inch
                        },
                    },
                },
                headers: {
                    default: new Header({
                        children: [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: "Chatbot Conversation Report",
                                        bold: true,
                                        size: 24,
                                    }),
                                ],
                                alignment: AlignmentType.CENTER,
                            }),
                        ],
                    }),
                },
                footers: {
                    default: new Footer({
                        children: [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `Generated on ${new Date().toLocaleString()}`,
                                        size: 18,
                                        color: "666666",
                                    }),
                                ],
                                alignment: AlignmentType.CENTER,
                            }),
                        ],
                    }),
                },
                children: [
                    // Title
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Chatbot Conversation Report",
                                bold: true,
                                size: 32,
                                color: "2E86AB",
                            }),
                        ],
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 400 },
                    }),
                    
                    // Metadata section
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Session Details",
                                bold: true,
                                size: 24,
                                underline: {},
                            }),
                        ],
                        spacing: { after: 200 },
                    }),
                    
                    new Paragraph({
                        children: [
                            new TextRun({ text: "User: ", bold: true, size: 22 }),
                            new TextRun({ text: "Guest", size: 22 }),
                        ],
                        spacing: { after: 100 },
                    }),
                    
                    new Paragraph({
                        children: [
                            new TextRun({ text: "Date: ", bold: true, size: 22 }),
                            new TextRun({ text: new Date().toLocaleDateString(), size: 22 }),
                        ],
                        spacing: { after: 100 },
                    }),
                    
                    new Paragraph({
                        children: [
                            new TextRun({ text: "Time: ", bold: true, size: 22 }),
                            new TextRun({ text: new Date().toLocaleTimeString(), size: 22 }),
                        ],
                        spacing: { after: 100 },
                    }),
                    
                    new Paragraph({
                        children: [
                            new TextRun({ text: "Total Questions: ", bold: true, size: 22 }),
                            new TextRun({ text: allQaStream.length.toString(), size: 22 }),
                        ],
                        spacing: { after: 400 },
                    }),
                    
                    // Conversation section
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Conversation",
                                bold: true,
                                size: 24,
                                underline: {},
                            }),
                        ],
                        spacing: { after: 300 },
                    }),
                    
                    // Q&A pairs
                    ...allQaStream.flatMap((qa, idx) => {
                        const elements = [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `Question ${idx + 1}:`,
                                        bold: true,
                                        size: 22,
                                        color: "2E86AB",
                                    }),
                                ],
                                spacing: { before: 200, after: 100 },
                            }),
                            
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: qa.question,
                                        size: 22,
                                    }),
                                ],
                                spacing: { after: 150 },
                                indent: { left: 360 }, // 0.25 inch indent
                            }),
                            
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `Answer ${idx + 1}:`,
                                        bold: true,
                                        size: 22,
                                        color: "A23B72",
                                    }),
                                ],
                                spacing: { after: 100 },
                            }),
                            
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: qa.answer ?? "No answer provided",
                                        size: 22,
                                        color: "444444",
                                    }),
                                ],
                                spacing: { after: 300 },
                                indent: { left: 360 }, // 0.25 inch indent
                            }),
                        ];
                        // Add a page break after every 3 Q&A pairs to avoid cramming
                        if (idx > 0 && (idx + 1) % 3 === 0) {
                            const pageBreakParagraph = new Paragraph({});
                            pageBreakParagraph.addChildElement(new PageBreak());
                            elements.push(pageBreakParagraph);
                        }
                        return elements;
                    }),
                ],
            },
        ],
    });
    
    const blob = await Packer.toBlob(doc);
    saveAs(blob, "chatbot-conversation-report.docx");
};

    const handleSendSelectedToChatbot = async () => {
        if (!results) return;
        setSending(true);
        setQaStream([]);
        setCurrentlyProcessing(null);

        // Gather selected paragraphs' content
        const selectedParaKeys = Object.keys(selectedParagraphs).filter(k => selectedParagraphs[k]);
        const selectedParagraphsContent = selectedParaKeys.map(key => {
            const [sectionIdx] = key.split('-').map(Number);
            const section = results.headerSections[sectionIdx];
            return `Section: ${section.title}\n${section.content}`;
        });

        // Send selected paragraphs as individual queries
        for (let idx = 0; idx < selectedParaKeys.length; idx++) {
            setCurrentlyProcessing(idx);
            const [sectionIdx, paraIdx] = selectedParaKeys[idx].split('-').map(Number);
            const section = results.headerSections[sectionIdx];
            const paragraph = section.paragraphs[paraIdx];
            const context = `Section: ${section.title}\nParagraph ${paragraph.paragraphNumber}: ${paragraph.content}`;
            try {
                const response = await chatApi({
                    messages: [{ content: context, role: "user" }],
                    session_state: {},
                }, false, undefined);
                const data = await response.json();
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
                setQaStream(prev => [
                    ...prev,
                    { question: `Section: ${section.title} - Paragraph ${paragraph.paragraphNumber}`, answer }
                ]);
            } catch (err) {
                setQaStream(prev => [
                    ...prev,
                    { question: `Section: ${section.title} - Paragraph ${paragraph.paragraphNumber}`, answer: "Error contacting chatbot." }
                ]);
            }
        }

        // Send selected questions, including selected paragraphs as context
        for (let idx = 0; idx < selectedQuestions.length; idx++) {
            setCurrentlyProcessing(selectedParaKeys.length + idx);
            const question = results.questions[selectedQuestions[idx]];
            // Combine all selected paragraphs as context for the question
            const context = `${selectedParagraphsContent.join('\n\n')}\n\nQuestion: ${question}`;
            try {
                const response = await chatApi({
                    messages: [{ content: context, role: "user" }],
                    session_state: {},
                }, false, undefined);
                const data = await response.json();
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
                setQaStream(prev => [
                    ...prev,
                    { question, answer }
                ]);
            } catch (err) {
                setQaStream(prev => [
                    ...prev,
                    { question, answer: "Error contacting chatbot." }
                ]);
            }
        }

        setCurrentlyProcessing(null);
        setSending(false);
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

    const processAllQuestions = async () => {
        if (!results) return;
        setSending(true);
        setQaStream([]);
        for (let idx = 0; idx < selectedQuestions.length; idx++) {
            setCurrentlyProcessing(idx);
            const qIdx = selectedQuestions[idx];
            const question = results.questions[qIdx];
            // Build context from selected titles and sections
            const context = createChatbotMessage({
                ...results,
                questions: [question], // Only this question
            });
            try {
                const response = await chatApi({
                    messages: [{ content: context, role: "user" }],
                    session_state: {},
                }, false, undefined);
                const data = await response.json();
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
                setQaStream(prev => [
                    ...prev,
                    { question, answer }
                ]);
            } catch (err) {
                setQaStream(prev => [
                    ...prev,
                    { question, answer: "Error contacting chatbot." }
                ]);
            }
        }
        setCurrentlyProcessing(null);
        setSending(false);
    };

    const processAllSections = async () => {
        if (!results) return;
        setSending(true);
        setQaStream([]);

        // Find the index of the "Scope" section (case-insensitive, may have number prefix)
        const scopeIdx = results.headerSections.findIndex(section =>
            /\bScope\b/i.test(section.title)
        );
        if (scopeIdx === -1) {
            setSending(false);
            return;
        }

        // Find the next section that starts with a number (e.g., "4", "5", etc.)
        let endIdx = results.headerSections.length;
        for (let i = scopeIdx + 1; i < results.headerSections.length; i++) {
            // Match headings like "4", "4.", "4 Title", "4. Title"
            if (/^\d+(\.| )/.test(results.headerSections[i].title.trim())) {
                endIdx = i;
                break;
            }
        }

        // Only process sections after "Scope" and before the next numbered section
        const sectionsToProcess = results.headerSections.slice(scopeIdx + 1, endIdx);

        for (let idx = 0; idx < sectionsToProcess.length; idx++) {
            setCurrentlyProcessing(idx);
            const section = sectionsToProcess[idx];
            // Build context: only the section heading and its content
            const context = `Section: ${section.title}\n\n${section.content}`;
            try {
                const response = await chatApi({
                    messages: [{ content: context, role: "user" }],
                    session_state: {},
                }, false, undefined);
                const data = await response.json();
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
                setQaStream(prev => [
                    ...prev,
                    { question: section.title, answer }
                ]);
            } catch (err) {
                setQaStream(prev => [
                    ...prev,
                    { question: section.title, answer: "Error contacting chatbot." }
                ]);
            }
        }
        setCurrentlyProcessing(null);
        setSending(false);
    };

    const processAllSectionsAndQuestions = async () => {
        if (!results) return;
        setSending(true);
        setQaStream([]);

        // Find the index of the "Scope" section (case-insensitive, may have number prefix)
        const scopeIdx = results.headerSections.findIndex(section =>
            /\bScope\b/i.test(section.title)
        );
        if (scopeIdx === -1) {
            setSending(false);
            return;
        }

        // Find the next section that starts with a number (e.g., "4", "5", etc.)
        let endIdx = results.headerSections.length;
        for (let i = scopeIdx + 1; i < results.headerSections.length; i++) {
            if (/^\d+(\.| )/.test(results.headerSections[i].title.trim())) {
                endIdx = i;
                break;
            }
        }

        // Only process sections after "Scope" and before the next numbered section
        const sectionsToProcess = results.headerSections.slice(scopeIdx + 1, endIdx);

        // Process sections
        for (let idx = 0; idx < sectionsToProcess.length; idx++) {
            setCurrentlyProcessing(idx);
            const section = sectionsToProcess[idx];
            const context = `Section: ${section.title}\n\n${section.content}`;
            try {
                const response = await chatApi({
                    messages: [{ content: context, role: "user" }],
                    session_state: {},
                }, false, undefined);
                const data = await response.json();
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
                setQaStream(prev => [
                    ...prev,
                    { question: section.title, answer }
                ]);
            } catch (err) {
                setQaStream(prev => [
                    ...prev,
                    { question: section.title, answer: "Error contacting chatbot." }
                ]);
            }
        }

        // Process questions as individual queries
        for (let qIdx = 0; qIdx < results.questions.length; qIdx++) {
            setCurrentlyProcessing(sectionsToProcess.length + qIdx);
            const question = results.questions[qIdx];
            // Optionally, you can include selected sections as context if needed
            const context = `Question: ${question}`;
            try {
                const response = await chatApi({
                    messages: [{ content: context, role: "user" }],
                    session_state: {},
                }, false, undefined);
                const data = await response.json();
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
                setQaStream(prev => [
                    ...prev,
                    { question, answer }
                ]);
            } catch (err) {
                setQaStream(prev => [
                    ...prev,
                    { question, answer: "Error contacting chatbot." }
                ]);
            }
        }

        setCurrentlyProcessing(null);
        setSending(false);
    };

    const processScopeAndTasks = async () => {
        if (!results) return;
        setSending(true);
        setQaStream([]);

        // Find the index and number of the "Scope" section (e.g., "3 Scope")
        const scopeIdx = results.headerSections.findIndex(section =>
            /\bScope\b/i.test(section.title)
        );
        if (scopeIdx === -1) {
            setSending(false);
            return;
        }

        // Extract the section number from the Scope title (e.g., "3 Scope" -> 3)
        const scopeTitle = results.headerSections[scopeIdx].title.trim();
        const scopeNumberMatch = scopeTitle.match(/^(\d+)/);
        const scopeNumber = scopeNumberMatch ? scopeNumberMatch[1] : null;

        // Collect all subsections under Scope (e.g., 3.1, 3.2, etc.)
        const sectionsToProcess = [results.headerSections[scopeIdx]];
        for (let i = scopeIdx + 1; i < results.headerSections.length; i++) {
            const title = results.headerSections[i].title.trim();
            // Match subsection like "3.1", "3.2", etc.
            if (scopeNumber && new RegExp(`^${scopeNumber}\\.`).test(title)) {
                sectionsToProcess.push(results.headerSections[i]);
            } else {
                // Stop at the next top-level section (e.g., "4", "5", etc.)
                if (/^\d+(\.| )/.test(title) && !new RegExp(`^${scopeNumber}\\.`).test(title)) {
                    break;
                }
            }
        }

        // Send each section (Scope and its tasks) as a separate query
        for (let idx = 0; idx < sectionsToProcess.length; idx++) {
            setCurrentlyProcessing(idx);
            const section = sectionsToProcess[idx];
            const context = `Section: ${section.title}\n\n${section.content}`;
            try {
                const response = await chatApi({
                    messages: [{ content: context, role: "user" }],
                    session_state: {},
                }, false, undefined);
                const data = await response.json();
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
                setQaStream(prev => [
                    ...prev,
                    { question: section.title, answer }
                ]);
            } catch (err) {
                setQaStream(prev => [
                    ...prev,
                    { question: section.title, answer: "Error contacting chatbot." }
                ]);
            }
        }

        // Send each question as a separate query, with only Scope section as context
        for (let qIdx = 0; qIdx < results.questions.length; qIdx++) {
            setCurrentlyProcessing(sectionsToProcess.length + qIdx);
            const question = results.questions[qIdx];
            const context = `Section: ${results.headerSections[scopeIdx].title}\n\n${results.headerSections[scopeIdx].content}\n\nQuestion: ${question}`;
            try {
                const response = await chatApi({
                    messages: [{ content: context, role: "user" }],
                    session_state: {},
                }, false, undefined);
                const data = await response.json();
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
                setQaStream(prev => [
                    ...prev,
                    { question, answer }
                ]);
            } catch (err) {
                setQaStream(prev => [
                    ...prev,
                    { question, answer: "Error contacting chatbot." }
                ]);
            }
        }

        setCurrentlyProcessing(null);
        setSending(false);
    };



    const handleAskNewQuestion = async () => {
        if (!newQuestion.trim()) return;
        setAddingQuestion(true);
        setCurrentlyProcessing(qaStream.length);
        // Use selected titles/sections as context for the new question
        const context = createChatbotMessage({
            ...results!,
            questions: [newQuestion]
        });
        try {
            const response = await chatApi({
                messages: [{ content: context, role: "user" }],
                session_state: {},
            }, false, undefined);
            const data = await response.json();
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
            setQaStream(prev => [
                ...prev,
                { question: newQuestion, answer }
            ]);
            setNewQuestion("");
        } catch (err) {
            setQaStream(prev => [
                ...prev,
                { question: newQuestion, answer: "Error contacting chatbot." }
            ]);
        }
        setCurrentlyProcessing(null);
        setAddingQuestion(false);
    };

    // Save current Q&A stream to local storage
    localStorage.setItem("rfpQaStream", JSON.stringify(qaStream));

    return (
        <div className={styles.container}>
            <div className={styles.header}>
                <h1>📄 RFI Processor</h1>
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
                    <div className={styles.results} style={{ display: "flex", gap: 32 }}>
                        {/* Left Sidebar: Extracted Content */}
                        <div
    className={styles.sidebar}
>

    {/* Header Sections with Content and Paragraphs */}
    <div className={styles.section}>
        <h3><span className={styles.sectionIcon}>📋</span>Document Sections with Paragraphs</h3>
        <div>
            {results.headerSections.length > 0
                ? results.headerSections.map((section, i) => (
                    <div className={styles.item} key={i} style={{ marginBottom: "20px" }}>
                        <div style={{ fontWeight: "600", marginBottom: "8px", color: "#2563eb" }}>
                            {section.title}
                        </div>
                        {section.paragraphs.map((paragraph, pIdx) => (
                            <div key={pIdx} style={{ display: "flex", alignItems: "center", marginBottom: "6px", marginLeft: 16 }}>
                                <input
                                    type="checkbox"
                                    checked={!!selectedParagraphs[`${i}-${pIdx}`]}
                                    onChange={() => {
                                        setSelectedParagraphs(prev => ({
                                            ...prev,
                                            [`${i}-${pIdx}`]: !prev[`${i}-${pIdx}`]
                                        }));
                                    }}
                                    style={{ marginRight: 8 }}
                                />
                                <span style={{ flex: 1, fontSize: "0.97em", color: "#444" }}>
                                    <strong>Paragraph {paragraph.paragraphNumber}:</strong> {paragraph.content}
                                </span>
                            </div>
                        ))}
                    </div>
                ))
                : <div className={styles.item}>No document sections found.</div>
            }
        </div>
    </div>

    {/* Questions Section */}
    <div className={styles.section}>
        <h3><span className={styles.sectionIcon}>❓</span>Questions</h3>
        <div>
            {results.questions.length > 0
                ? results.questions.map((question, i) => (
                    <div className={styles.item} key={i} style={{ display: "flex", alignItems: "center", marginBottom: "10px" }}>
                        <input
                            type="checkbox"
                            checked={selectedQuestions.includes(i)}
                            onChange={() => toggleQuestion(i)}
                            style={{ marginRight: 8 }}
                        />
                        <span style={{ flex: 1, fontWeight: "500", color: "#1e40af" }}>{question}</span>
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
        {/* Move the button here */}
        <button
            className={styles.copyBtn}
            style={{ marginTop: 24, width: "100%" }}
            disabled={sending || (selectedHeaderSections.length === 0 && selectedQuestions.length === 0)}
            onClick={handleSendSelectedToChatbot}
        >
            {sending ? "Sending..." : "Send Selected to Chatbot"}
        </button>
    </div>
</div>

                        {/* Right: Chatbot Q&A Stream */}
                        <div
    className={styles.qaStream}
   
>
    <h3 style={{ marginBottom: 24 }}>🤖 Chatbot Answers</h3>
    <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 18 }}>
        {qaStream.length === 0 && (
            <div style={{ color: "#888", textAlign: "center", marginTop: 60 }}>
                <span>No answers yet. Questions will appear here as they are processed.</span>
            </div>
        )}
        {qaStream.map((qa, idx) => (
            <div key={idx} style={{
                display: "flex",
                flexDirection: "column",
                alignItems: "flex-end",
                marginBottom: 8
            }}>
                <div style={{
                    background: "#2563eb",
                    color: "#fff",
                    borderRadius: "16px 16px 0 16px",
                    padding: "12px 18px",
                    maxWidth: "80%",
                    fontWeight: 500,
                    boxShadow: "0 2px 8px rgba(0,0,0,0.06)"
                }}>
                    <span style={{ fontSize: "0.95em", opacity: 0.8 }}>Q{idx + 1}: {qa.question}</span>
                </div>
                <div style={{
                    background: "#fff",
                    color: "#176317",
                    borderRadius: "0 16px 16px 16px",
                    padding: "12px 18px",
                    marginTop: 6,
                    maxWidth: "80%",
                    fontSize: "1em",
                    boxShadow: "0 2px 8px rgba(0,0,0,0.04)",
                    border: "1px solid #e0e7ff"
                }}>
                    {qa.answer
                        ? qa.answer
                        : (currentlyProcessing === idx ? "Processing..." : "Waiting...")}
                </div>
            </div>
        ))}
        {currentlyProcessing !== null && (
            <div style={{ marginTop: 12, color: "#888", textAlign: "center" }}>
                Processing item {currentlyProcessing + 1} of {Object.values(selectedParagraphs).filter(Boolean).length + selectedQuestions.length}...
            </div>
        )}
    </div>
    {/* Download buttons */}
    <div style={{ marginTop: 24, display: "flex", gap: 12, justifyContent: "flex-end" }}>
        <button className={styles.copyBtn} onClick={handleDownloadPdf}>📄 Download PDF</button>
        <button className={styles.copyBtn} onClick={handleDownloadWord}>📝 Download Word</button>
    </div>
    {/* Ask more questions */}
    <div style={{ marginTop: 32, display: "flex", gap: 8 }}>
        <input
            type="text"
            value={newQuestion}
            onChange={e => setNewQuestion(e.target.value)}
            placeholder="Ask another question..."
            style={{
                flex: 1,
                padding: "10px 14px",
                borderRadius: 8,
                border: "1px solid #d1d5db",
                fontSize: "1rem"
            }}
            disabled={addingQuestion}
            onKeyDown={e => {
                if (e.key === "Enter" && !addingQuestion && newQuestion.trim()) {
                    handleAskNewQuestion();
                }
            }}
        />
        <button
            className={styles.copyBtn}
            onClick={handleAskNewQuestion}
            disabled={addingQuestion || !newQuestion.trim()}
            style={{ minWidth: 120 }}
        >
            {addingQuestion ? "Asking..." : "Ask"}
        </button>
    </div>
</div>
                    </div>
                )}
            </div>
        </div>
    );
};
