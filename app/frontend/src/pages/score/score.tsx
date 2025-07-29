import React, { useEffect, useState } from "react";
import { chatApi } from "../../api/api"; // Adjust path as needed
import styles from "./score.module.css"; // Create this CSS file for custom styles
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { Document, Packer, Paragraph, TextRun } from "docx";
import { saveAs } from "file-saver";

// List of programs to score
const programs = [
    "NGA NSG/SI",
    "NGA ESMARTS",
    "NGA RDAS",
    "NGA SEIN",
    "NGA SEIN-ASI-SB",
    "ODNI CASES",
    "NGA EMERALD",
    "NGA MOJAVE",
    "NGA NSE",
    "NRO COLOSSUS",
    "NCTC TPI",
    "NGA WSTAMP",
    "NGA RTSS",
    "NGA S3",
    "USA AI/ML Facial Recognition",
    "NRO LANDMARK AOS",
    "NRO ISPO",
    "NGA ARSO"
];

// Chatbot scoring function that makes actual API calls
async function getScoreFromChatbot(program: string, topic: string): Promise<{score: number, reason: string}> {
    try {
        // First, ask the context-setting question
        const contextPrompt = `Where program = ${program}, rate the relevance to the topic is ${topic}?`;
        await chatApi({
            messages: [{ content: contextPrompt, role: "user" }],
            session_state: {}
        }, false, undefined);

        // Then, ask the scoring question
        const scoringPrompt = `Score the relevance of the following program to the given topic:

Program: "${program}"
Topic: "${topic}"

USE this exact structure:

Score: X/5
Reason: [Detailed explanation]

Use this scoring framework:
- 0/5: Unrelated or incompatible
- 1/5: Minimal relevance 
- 2/5: Limited relevance with some indirect connections
- 3/5: Moderate relevance 
- 4/5: High relevance with strong connections and alignment
- 5/5: Perfect match 

Provide your assessment now.`;

        const response = await chatApi({
            messages: [{ content: scoringPrompt, role: "user" }],
            session_state: {}
        }, false, undefined);

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        const answer = data.answer?.content ||
                       data.answer ||
                       data.message?.content ||
                       data.message ||
                       data.choices?.[0]?.message?.content ||
                       data.response ||
                       "";

        // Parse the response to extract score and reason
        const scoreMatches = answer.match(/score[:\s]*([0-5])\/?5?/gi);
        let score: number | null = null;
        if (scoreMatches && scoreMatches.length > 0) {
            // Get the last match and extract the digit
            const lastScoreMatch = scoreMatches[scoreMatches.length - 1];
            const digitMatch = lastScoreMatch.match(/([0-5])\/?5?/);
            if (digitMatch) {
                score = parseInt(digitMatch[1], 10);
            }
        }
        const reasonMatch = answer.match(/reason[:\s]*(.*?)(?:\n|$)/is);

        if (score != null) {
            const reason = reasonMatch ? reasonMatch[1].trim() : "No detailed reason provided.";
            return { score, reason };
        }

        // Fallback parsing - look for just numbers
        const numberMatch = answer.match(/([1-5])/);
        if (numberMatch) {
            return {
                score: parseInt(numberMatch[1], 10),
                reason: answer.replace(/score[:\s]*[1-5]\/?5?/i, '').replace(/reason[:\s]*/i, '').trim() || "Score extracted from response."
            };
        }

        // If no score found, return the full response as reason with default score
        return {
            score: 1,
            reason: answer || "No valid response received from chatbot."
        };

    } catch (error) {
        console.error("Error calling chatbot API:", error);
        return {
            score: 1,
            reason: `Error occurred while getting score: ${typeof error === "object" && error !== null && "message" in error ? (error as { message: string }).message : String(error)}`
        };
    }
}

export function Score() {
    const [topic, setTopic] = useState("");
    const [results, setResults] = useState<{ [program: string]: {score: number, reason: string} | null }>({});
    const [currentIdx, setCurrentIdx] = useState<number>(0);
    const [processing, setProcessing] = useState<boolean>(false);
    const [started, setStarted] = useState<boolean>(false);
    const [customPrograms, setCustomPrograms] = useState<string[] | null>(null);
    const [selectedPrograms, setSelectedPrograms] = useState<{ [program: string]: boolean }>({});
    const [fullscreenProgram, setFullscreenProgram] = useState<string | null>(null);
    const [defaultPrograms, setDefaultPrograms] = useState<string[] | null>(null);
    const [userName, setUserName] = useState("User");

    const programList = defaultPrograms ?? programs;

    // Filter programs based on selectedPrograms (default to true if not set)
    const filteredProgramList = programList.filter(p => selectedPrograms[p] ?? true);

    useEffect(() => {
        const stored = localStorage.getItem("defaultPrograms");
        if (stored) {
            try {
                const parsed = JSON.parse(stored);
                if (Array.isArray(parsed) && parsed.every(x => typeof x === "string")) {
                    setDefaultPrograms(parsed);
                }
            } catch {}
        }
    }, []);

    useEffect(() => {
        if (started && !processing && currentIdx < filteredProgramList.length) {
            setProcessing(true);
            getScoreFromChatbot(filteredProgramList[currentIdx], topic)
                .then(({score, reason}) => {
                    setResults(prev => ({
                        ...prev,
                        [filteredProgramList[currentIdx]]: { score, reason }
                    }));
                    setCurrentIdx(idx => idx + 1);
                    setProcessing(false);
                })
                .catch(error => {
                    setResults(prev => ({
                        ...prev,
                        [filteredProgramList[currentIdx]]: { 
                            score: 1, 
                            reason: `Error occurred while scoring: ${error.message}` 
                        }
                    }));
                    setCurrentIdx(idx => idx + 1);
                    setProcessing(false);
                });
        }
    }, [started, currentIdx, processing, topic, filteredProgramList]);

    // const programList = customPrograms ?? programs; // Removed duplicate declaration

    const handleStart = () => {
        if (topic.trim()) {
            setResults({});
            setCurrentIdx(0);
            setStarted(true);
        }
    };

    const handleReset = () => {
        setResults({});
        setCurrentIdx(0);
        setStarted(false);
        setProcessing(false);
    };

    const handleProgramUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;
        const arrayBuffer = await file.arrayBuffer();
        const mammoth = await import("mammoth");
        const { value: text } = await mammoth.extractRawText({ arrayBuffer });
        const names = text.split('\n').map(line => line.trim()).filter(Boolean);
        if (names.length > 0) {
            setDefaultPrograms(names);
            localStorage.setItem("defaultPrograms", JSON.stringify(names)); // <-- Save to localStorage
        } else {
            alert("No program names found in the uploaded file.");
        }
    };

    const handleDownloadDefaultPrograms = async () => {
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({
                        children: [new TextRun({ text: "Default Program List", bold: true, size: 32 })],
                        spacing: { after: 400 }
                    }),
                    ...programs.map(p => new Paragraph({ children: [new TextRun({ text: p, size: 24 })], spacing: { after: 100 } }))
                ]
            }]
        });
        const blob = await Packer.toBlob(doc);
        saveAs(blob, "default-programs.docx");
    };
    const completedCount = filteredProgramList.filter(p => results[p] != null).length;
    const progressPercentage = (completedCount / filteredProgramList.length) * 100;

    const getScoreColor = (score: number) => {
        if (score >= 4) return '#10b981'; // green
        if (score >= 3) return '#f59e0b'; // yellow
        return '#ef4444'; // red
    };

    // Only include selected programs
    const selectedProgramNames = programList.filter(p => selectedPrograms[p] ?? true);

    // PDF export
    const handleExportPDF = () => {
        const doc = new jsPDF();
        const now = new Date();
        const dateStr = now.toLocaleDateString();
        const timeStr = now.toLocaleTimeString();

        doc.setFontSize(18);
        doc.text("Program Relevancy Scorer Results", 14, 18);
        doc.setFontSize(12);
        doc.text(`User: ${userName}`, 14, 26);
        doc.text(`Date: ${dateStr}    Time: ${timeStr}`, 14, 34);
        doc.text(`Topic: ${topic}`, 14, 42); // Added topic to PDF

        autoTable(doc, {
            startY: 50, // was 42, now after topic
            head: [["Program", "Score", "Reason"]],
            body: selectedProgramNames.map(p => [
                p,
                results[p]?.score ?? "-",
                results[p]?.reason ?? "-"
            ]),
            styles: { fontSize: 10, cellWidth: 'wrap' },
            columnStyles: { 2: { cellWidth: 100 } }
        });
        doc.save("program-scores.pdf");
    };

    // Word export
    const handleExportWord = async () => {
        const now = new Date();
        const dateStr = now.toLocaleDateString();
        const timeStr = now.toLocaleTimeString();

        const paragraphs = [
            new Paragraph({ children: [new TextRun({ text: "Program Relevancy Scorer Results", bold: true, size: 32 })], spacing: { after: 400 } }),
            new Paragraph({ children: [new TextRun({ text: `User: ${userName}`, italics: true, size: 24 })] }),
            new Paragraph({ children: [new TextRun({ text: `Date: ${dateStr}    Time: ${timeStr}`, italics: true, size: 24 })] }),
            new Paragraph({ children: [new TextRun({ text: `Topic: ${topic}`, bold: true, size: 24 })] }),
            new Paragraph({})
        ];

        selectedProgramNames.forEach(p => {
            paragraphs.push(
                new Paragraph({
                    children: [
                        new TextRun({ text: p, bold: true, size: 28, color: "be4c29" }),
                        new TextRun({ text: `  Score: ${results[p]?.score ?? "-"}/5`, bold: true, size: 24 })
                    ],
                    spacing: { after: 100 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({ text: results[p]?.reason ?? "-", size: 22 })
                    ],
                    spacing: { after: 300 }
                })
            );
        });

        const doc = new Document({
            sections: [{ children: paragraphs }]
        });
        const blob = await Packer.toBlob(doc);
        saveAs(blob, "program-scores.docx");
    };



    return (
        <div className={styles.container}>
            <h2 className={styles.title}>Program Relevancy Scorer</h2>

            <div className={styles.inputContainer}>
                <input
                    type="text"
                    value={topic}
                    onChange={e => setTopic(e.target.value)}
                    placeholder="Enter a topic (e.g., cloud computing, AI, geospatial analysis)"
                    className={styles.topicInput}
                    disabled={started && currentIdx < programList.length}
                    onKeyDown={e => {
                        if (e.key === "Enter" && topic.trim() && !(started && currentIdx < programList.length)) {
                            handleStart();
                        }
                    }}
                />
                <button
                    onClick={handleStart}
                    disabled={!topic.trim() || (started && currentIdx < programList.length)}
                    className={styles.startButton}
                >
                    {started && currentIdx < programs.length ? "Scoring..." : "Start Scoring"}
                </button>
                {started && (
                    <button
                        onClick={handleReset}
                        className={styles.startButton}
                        style={{
                            background: '#6b7280'
                        }}
                    >
                        Reset
                    </button>
                )}
            </div>

            <div style={{ display: "flex", gap: 16, marginBottom: 24, alignItems: "center" }}>
                <label>
                    <input
                        type="file"
                        accept=".docx"
                        style={{ display: "none" }}
                        onChange={handleProgramUpload}
                    />
                    <button
                        className={styles.startButton}
                        type="button"
                        onClick={() => document.querySelector<HTMLInputElement>('input[type="file"]')?.click()}
                    >
                        Upload Program List (.docx)
                    </button>
                </label>
                <button className={styles.startButton} type="button" onClick={handleExportPDF}>
                    Print to PDF
                </button>
                <button className={styles.startButton} type="button" onClick={handleExportWord}>
                    Print to Word
                </button>
            </div>

            {started && (
                <div className={styles.progressContainer}>
                    <div className={styles.progressInfo}>
                        <span className={styles.progressText}>
                            Progress: {completedCount} of {filteredProgramList.length} programs
                        </span>
                        <span className={styles.progressPercentage}>
                            {Math.round(progressPercentage)}%
                        </span>
                    </div>
                    <div className={styles.progressBar}>
                        <div
                            className={styles.progressFill}
                            style={{ width: `${progressPercentage}%` }}
                        />
                    </div>
                </div>
            )}

            {/* Fullscreen program view */}
            {fullscreenProgram && (
                <div className={styles.fullscreenOverlay}>
                    <div className={styles.fullscreenContent}>
                        <button
                            className={styles.closeButton}
                            onClick={() => setFullscreenProgram(null)}
                        >
                            ×
                        </button>
                        <h3 className={styles.programName}>{fullscreenProgram}</h3>
                        {results[fullscreenProgram] ? (
                            <>
                                <div
                                    className={styles.chatbotScore}
                                    style={{ color: getScoreColor(results[fullscreenProgram]?.score || 1) }}
                                >
                                    Score: {results[fullscreenProgram]?.score}/5
                                </div>
                                <div className={styles.reasonBox}>
                                    <span className={styles.reasonLabel}>Reason:</span> {results[fullscreenProgram]?.reason}
                                </div>
                            </>
                        ) : (
                            <span className={styles.status}>No result yet.</span>
                        )}
                    </div>
                </div>
            )}

            <ul className={styles.programList}>
                {(started ? filteredProgramList : programList).map((program, idx) => {
                    const checked = selectedPrograms[program] ?? true;
                    return (
                        <li
                            className={styles.programItem}
                            key={program}
                            onClick={() => {
                                if (started && results[program]) {
                                    setFullscreenProgram(program);
                                } else if (!started) {
                                    setSelectedPrograms(prev => ({
                                        ...prev,
                                        [program]: !checked
                                    }));
                                }
                            }}
                            style={{ cursor: started && results[program] ? 'pointer' : !started ? 'pointer' : undefined }}
                        >
                            {!started && (
                                <input
                                    type="checkbox"
                                    checked={checked}
                                    onChange={e => {
                                        setSelectedPrograms(prev => ({
                                            ...prev,
                                            [program]: e.target.checked
                                        }));
                                    }}
                                    style={{ marginRight: 12, pointerEvents: 'none' }}
                                />
                            )}
                            <span className={styles.programName}>{program}</span>
                            <div className={styles.resultContent}>
                                {started ? (
                                    results[program] != null ? (
                                        <>
                                            <div
                                                className={styles.chatbotScore}
                                                style={{ color: getScoreColor(results[program]?.score || 1) }}
                                            >
                                                Score: {results[program]?.score}/5
                                            </div>
                                            <div className={styles.reasonBox}>
                                                <span className={styles.reasonLabel}>Reason:</span> {results[program]?.reason}
                                            </div>
                                        </>
                                    ) : (
                                        <span className={`${styles.status} ${started && idx === currentIdx ? styles.statusActive : ''}`}>
                                            {started && idx === currentIdx ? (
                                                <span>
                                                    <span className={styles.loadingDot}>●</span> Scoring...
                                                </span>
                                            ) : (
                                                "Waiting..."
                                            )}
                                        </span>
                                    )
                                ) : null}
                            </div>
                        </li>
                    );
                })}
            </ul>
            
            {started && currentIdx >= filteredProgramList.length && (
                <div className="done">
                    ✅ All programs scored for topic: <strong>"{topic}"</strong>
                </div>
            )}
        </div>
    );
}