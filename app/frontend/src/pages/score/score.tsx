import React, { useState, useEffect } from "react";
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
        const prompt = `Rate the relevance of the program "${program}" to the topic "${topic}" on a scale of 1-5, where:
1 = Not relevant at all
2 = Slightly relevant
3 = Moderately relevant
4 = Highly relevant
5 = Extremely relevant

Please respond in EXACTLY this format:
Score: X/5
Reason: [Your detailed explanation of why this score was given]

Consider the program's purpose, capabilities, and how it might relate to or benefit from the given topic.`;

        const response = await chatApi({
            messages: [{ content: prompt, role: "user" }],
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

        console.log("Chatbot answer:", answer);

        // Parse the response to extract score and reason
        const scoreMatch = answer.match(/score[:\s]*([1-5])\/?5?/i);
        const reasonMatch = answer.match(/reason[:\s]*(.*?)(?:\n|$)/is);

        if (scoreMatch) {
            const score = parseInt(scoreMatch[1], 10);
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

    const programList = customPrograms ?? programs; // Removed duplicate declaration

    // Filter programs based on selectedPrograms (default to true if not set)
    const filteredProgramList = programList.filter(p => selectedPrograms[p] ?? true);

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
        // Use mammoth to extract text from docx file
        // Install mammoth: npm install mammoth
        const mammoth = await import("mammoth");
        const { value: text } = await mammoth.extractRawText({ arrayBuffer });
        // Split lines and filter out empty lines
        const names = text.split('\n').map(line => line.trim()).filter(Boolean);
        if (names.length > 0) setCustomPrograms(names);
        else alert("No program names found in the uploaded file.");
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

    // PDF export
    const handleExportPDF = () => {
        const doc = new jsPDF();
        doc.setFontSize(18);
        doc.text("Program Relevancy Scorer Results", 14, 18);
        autoTable(doc, {
            startY: 28,
            head: [["Program", "Score", "Reason"]],
            body: programs.map(p => [
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
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({
                        children: [new TextRun({ text: "Program Relevancy Scorer Results", bold: true, size: 32 })],
                        spacing: { after: 400 }
                    }),
                    ...programs.map(p =>
                        [
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
                        ]
                    ).flat()
                ]
            }]
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
                <button
                    className={styles.startButton}
                    type="button"
                    onClick={handleDownloadDefaultPrograms}
                >
                    Download Default Program List
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

            <ul className={styles.programList}>
                {(started ? filteredProgramList : programList).map((program, idx) => (
                    <li className={styles.programItem} key={program}>
                        {!started && (
                            <input
                                type="checkbox"
                                checked={selectedPrograms[program] ?? true}
                                onChange={e =>
                                    setSelectedPrograms(prev => ({
                                        ...prev,
                                        [program]: e.target.checked
                                    }))
                                }
                                style={{ marginRight: 12 }}
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
                ))}
            </ul>
            
            {started && currentIdx >= filteredProgramList.length && (
                <div className="done">
                    ✅ All programs scored for topic: <strong>"{topic}"</strong>
                </div>
            )}
        </div>
    );
}