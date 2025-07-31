import React, { useEffect, useState } from "react";
import { chatApi } from "../../api/api"; // Adjust path as needed
import styles from "./score.module.css"; // Create this CSS file for custom styles
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { Document, Packer, Paragraph, TextRun } from "docx";
import { saveAs } from "file-saver";

// Default hierarchical programs and task orders
const defaultProgramHierarchy: { [program: string]: string[] } = {
    "NGA NSG/SI": [],
    "NGA ESMARTS": [
        "TO01 X4",
        "TO02 XK",
        "TO03 Enterprise Operations",
        "TO04 OCIO CS"
    ],
    "NGA RDAS": [
        "TO04 Innovision",
        "TO11 International",
        "TO13 ITEMS",
        "TO17 Office of GEOINT Management",
        "TO18 Tactical Data Program Management",
        "TO21 ITEMS PMO",
        "TO24 NSG Expeditionary Architecture",
        "TO26 CAE",
        "TO27 Online GEOINT Services"
    ],
    "NGA SEIN": [],
    "NGA SEIN-ASI-SB": [],
    "ODNI CASES": [
        "TO01 AT/CCT"
    ],
    "NGA EMERALD": [
        "TO02 National Technical Means",
        "TO06 Source Content Conveyance",
        "TO08 Research",
        "TO17 Source",
        "TO19 N2W",
        "TO20 Office of Content Solutions",
        "TO23 GEOINT Services",
        "TO30 GEOINT Enterprise",
        "TO31 Open IT Solutions",
        "TO32 IC Enterprise Management",
        "TO34 ITEMS and IC ITE",
        "TO43 OVI",
        "TO48 IPF"
    ],
    "NGA MOJAVE": [
        "TO02 RFO",
        "TO04 ATP",
        "TO12 SI"
    ],
    "NGA NSE": [
        "TO05 IPA"
    ],
    "NRO COLOSSUS": [],
    "NCTC TPI": [],
    "NGA WSTAMP": [],
    "NGA RTSS": [],
    "NGA S3": [],
    "USA AI/ML Facial Recognition": [],
    "NRO LANDMARK AOS": [],
    "NRO ISPO": [],
    "NGA ARSO": []
};

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

// --- Update getScoreFromChatbot to support taskOrder ---
async function getScoreFromChatbot(
    program: string,
    topic: string,
    selectedIndex: string,
    taskOrder?: string // <-- new param
): Promise<{ score: number, reason: string }> {
    try {
        // Build context and scoring prompts based on taskOrder
        let contextPrompt: string;
        let scoringPrompt: string;

        if (taskOrder) {
            contextPrompt = `Where program = ${program} and taskOrder = ${taskOrder}, rate the relevance to the topic is ${topic}?`;
            scoringPrompt = `Score the relevance of the following program and task order to the given topic:

Program: "${program}"
Task Order: "${taskOrder}"
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
        } else {
            contextPrompt = `Where program = ${program}, rate the relevance to the topic is ${topic}?`;
            scoringPrompt = `Score the relevance of the following program to the given topic:

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
        }

        await chatApi({
            messages: [{ content: contextPrompt, role: "user" }],
            session_state: {}
        }, false, undefined);

        const response = await chatApi({
            messages: [{ content: scoringPrompt, role: "user" }],
            session_state: {},
            context: {
                overrides: {
                    search_index: selectedIndex,
                    vector_fields: "",
                    language: "en",
                    use_agentic_retrieval: false
                }
            }
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
    const [fullscreenProgram, setFullscreenProgram] = useState<string | null>(null);
    const [defaultPrograms, setDefaultPrograms] = useState<string[] | null>(() => Object.keys(defaultProgramHierarchy));
    const [userName, setUserName] = useState("User");
    const [indexes, setIndexes] = useState<string[]>([]);
    const [selectedIndex, setSelectedIndex] = useState<string>("");
    const [showIndexSwitcher, setShowIndexSwitcher] = useState(false);
    const [pendingIndex, setPendingIndex] = useState(selectedIndex);
    const [hierarchicalPrograms, setHierarchicalPrograms] = useState<{ [program: string]: string[] }>(() => {
        // On first load, check localStorage, else use default
        const storedHierarchy = localStorage.getItem("defaultProgramsHierarchy");
        if (storedHierarchy) {
            try {
                const parsed = JSON.parse(storedHierarchy);
                if (parsed && typeof parsed === "object") return parsed;
            } catch {}
        }
        return defaultProgramHierarchy;
    });
    const [selectedPrograms, setSelectedPrograms] = useState<{ [program: string]: boolean }>({});
    const [selectedTaskOrders, setSelectedTaskOrders] = useState<{ [program: string]: string | null }>({});
    const [showProgramPopup, setShowProgramPopup] = useState<null | string>(null);

    // Helper to get all programs (from hierarchicalPrograms if present, else fallback)
    const programList = React.useMemo(() => {
        if (Object.keys(hierarchicalPrograms).length > 0) {
            return Object.keys(hierarchicalPrograms);
        }
        return defaultPrograms ?? programs;
    }, [hierarchicalPrograms, defaultPrograms]);

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

    // On mount, always load hierarchy from localStorage if present, else use default
    useEffect(() => {
        const storedHierarchy = localStorage.getItem("defaultProgramsHierarchy");
        if (storedHierarchy) {
            try {
                const parsed = JSON.parse(storedHierarchy);
                if (parsed && typeof parsed === "object") {
                    setHierarchicalPrograms(parsed);
                    setDefaultPrograms(Object.keys(parsed));
                    // Set all programs checked by default
                    const newSelectedPrograms: { [program: string]: boolean } = {};
                    Object.keys(parsed).forEach(p => { newSelectedPrograms[p] = true; });
                    setSelectedPrograms(newSelectedPrograms);
                    setSelectedTaskOrders({});
                    return;
                }
            } catch {}
        }
        // If nothing in storage, use the default
        setHierarchicalPrograms(defaultProgramHierarchy);
        setDefaultPrograms(Object.keys(defaultProgramHierarchy));
        const newSelectedPrograms: { [program: string]: boolean } = {};
        Object.keys(defaultProgramHierarchy).forEach(p => { newSelectedPrograms[p] = true; });
        setSelectedPrograms(newSelectedPrograms);
        setSelectedTaskOrders({});
    }, []);

    useEffect(() => {
        if (started && !processing && currentIdx < filteredProgramList.length) {
            setProcessing(true);
            const program = filteredProgramList[currentIdx];
            const taskOrder = selectedTaskOrders[program] || undefined;
            getScoreFromChatbot(program, topic, selectedIndex, taskOrder)
                .then(({ score, reason }) => {
                    setResults(prev => ({
                        ...prev,
                        [program]: { score, reason }
                    }));
                    setCurrentIdx(idx => idx + 1);
                    setProcessing(false);
                })
                .catch(error => {
                    setResults(prev => ({
                        ...prev,
                        [program]: {
                            score: 1,
                            reason: `Error occurred while scoring: ${error.message}`
                        }
                    }));
                    setCurrentIdx(idx => idx + 1);
                    setProcessing(false);
                });
        }
    }, [started, currentIdx, processing, topic, filteredProgramList, selectedIndex, selectedTaskOrders]);

    useEffect(() => {
        fetch("/api/search-indexes")
            .then(res => res.json())
            .then(data => {
                if (data.indexes && Array.isArray(data.indexes)) {
                    setIndexes(data.indexes);
                    setSelectedIndex(data.indexes[0] || "");
                }
            })
            .catch(err => console.error("Error fetching indexes:", err));
    }, []);

    // --- UPDATED: handleProgramUpload ---
    const handleProgramUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;
        const arrayBuffer = await file.arrayBuffer();
        const mammoth = await import("mammoth");
        const { value: text } = await mammoth.extractRawText({ arrayBuffer });
        const lines = text.split('\n').map(line => line.trim()).filter(Boolean);

        // Parse hierarchical structure
        const hierarchy: { [program: string]: string[] } = {};
        let currentProgram: string | null = null;
        for (const line of lines) {
            if (line.startsWith('-')) {
                if (currentProgram) {
                    const taskOrder = line.replace(/^-+/, '').trim();
                    if (taskOrder) {
                        hierarchy[currentProgram].push(taskOrder);
                    }
                }
            } else {
                currentProgram = line;
                hierarchy[currentProgram] = [];
            }
        }

        setHierarchicalPrograms(hierarchy);
        setDefaultPrograms(Object.keys(hierarchy));
        // Set all programs checked by default
        const newSelectedPrograms: { [program: string]: boolean } = {};
        Object.keys(hierarchy).forEach(p => { newSelectedPrograms[p] = true; });
        setSelectedPrograms(newSelectedPrograms);
        setSelectedTaskOrders({});
        // Overwrite the full hierarchy as the new default in localStorage
        localStorage.setItem("defaultProgramsHierarchy", JSON.stringify(hierarchy));
    };

    // Always sync selectedPrograms with the current program list
    useEffect(() => {
        const allPrograms = Object.keys(hierarchicalPrograms).length > 0
            ? Object.keys(hierarchicalPrograms)
            : (defaultPrograms ?? programs);

        setSelectedPrograms(prev => {
            const updated: { [program: string]: boolean } = {};
            allPrograms.forEach(p => {
                updated[p] = prev[p] !== undefined ? prev[p] : true;
            });
            return updated;
        });
    }, [hierarchicalPrograms, defaultPrograms]);

    useEffect(() => {
        if (started && !processing && currentIdx < filteredProgramList.length) {
            setProcessing(true);
            const program = filteredProgramList[currentIdx];
            const taskOrder = selectedTaskOrders[program] || undefined;
            getScoreFromChatbot(program, topic, selectedIndex, taskOrder)
                .then(({ score, reason }) => {
                    setResults(prev => ({
                        ...prev,
                        [program]: { score, reason }
                    }));
                    setCurrentIdx(idx => idx + 1);
                    setProcessing(false);
                })
                .catch(error => {
                    setResults(prev => ({
                        ...prev,
                        [program]: {
                            score: 1,
                            reason: `Error occurred while scoring: ${error.message}`
                        }
                    }));
                    setCurrentIdx(idx => idx + 1);
                    setProcessing(false);
                });
        }
    }, [started, currentIdx, processing, topic, filteredProgramList, selectedIndex, selectedTaskOrders]);

    // --- UPDATED: Checkbox logic for programs and task orders ---
    // When a program is unchecked, also uncheck its selected task order
    const handleProgramCheckbox = (program: string, checked: boolean) => {
        setSelectedPrograms(prev => ({
            ...prev,
            [program]: checked
        }));
        if (!checked) {
            setSelectedTaskOrders(prev => ({
                ...prev,
                [program]: null
            }));
        }
    };

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

    // --- Replace handleDownloadDefaultPrograms to export the hierarchical structure --- 
    const handleDownloadDefaultPrograms = async () => {
        // Export the current default program/task order hierarchy as a .docx in the same format as upload
        // Example:
        // NGA NSG/SI
        // NGA ESMARTS
        // - TO01 X4
        // - TO02 XK
        // ...
        let lines: string[] = [];
        const hierarchy = hierarchicalPrograms && Object.keys(hierarchicalPrograms).length > 0
            ? hierarchicalPrograms
            : defaultProgramHierarchy;

        Object.entries(hierarchy).forEach(([program, tasks]) => {
            lines.push(program);
            tasks.forEach(task => {
                lines.push(`- ${task}`);
            });
        });

        const paragraphs = lines.map(line =>
            new Paragraph({
                children: [new TextRun({ text: line, size: 24 })],
                spacing: { after: 100 }
            })
        );

        const doc = new Document({
            sections: [{
                children: [
                    //new Paragraph({
                    //    children: [new TextRun({ text: "Default Program List", bold: true, size: 32 })],
                    //    spacing: { after: 400 }
                    //}),
                    ...paragraphs
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
        const now = new Date();
        const dateStr = now.toLocaleDateString();
        const timeStr = now.toLocaleTimeString();

        doc.setFontSize(18);
        doc.text("Program Relevancy Scorer Results", 14, 18);
        doc.setFontSize(12);
        doc.text(`User: ${userName}`, 14, 26);
        doc.text(`Date: ${dateStr}    Time: ${timeStr}`, 14, 34);
        doc.text(`Topic: ${topic}`, 14, 42);

        autoTable(doc, {
            startY: 50,
            head: [["Program", "Task Order", "Score", "Reason"]],
            body: filteredProgramList.map(p => [
                p,
                selectedTaskOrders[p] || "-",
                results[p]?.score ?? "-",
                results[p]?.reason ?? "-"
            ]),
            styles: { fontSize: 10, cellWidth: 'wrap' },
            columnStyles: { 3: { cellWidth: 100 } }
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

        filteredProgramList.forEach(p => {
            paragraphs.push(
                new Paragraph({
                    children: [
                        new TextRun({ text: p, bold: true, size: 28, color: "be4c29" }),
                        ...(selectedTaskOrders[p]
                            ? [new TextRun({ text: `  Task Order: ${selectedTaskOrders[p]}`, bold: true, size: 24, color: "764ba2" })]
                            : []),
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
                <button
                    className={styles.startButton}
                    type="button"
                    onClick={() => {
                        setPendingIndex(selectedIndex);
                        setShowIndexSwitcher(true);
                    }}
                    disabled={started && currentIdx < programList.length}
                >
                    Switch Index
                </button>
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
                    Download Current Default List
                </button>
                <button className={styles.startButton} type="button" onClick={handleExportPDF}>
                    Print to PDF
                </button>
                <button className={styles.startButton} type="button" onClick={handleExportWord}>
                    Print to Word
                </button>
            </div>
            <div style={{ marginBottom: 16, fontWeight: 500, color: "#764ba2" }}>
                Current Microsoft Index: <span style={{ fontWeight: 700 }}>{selectedIndex}</span>
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

            {/* Program selection popup */}
            {showProgramPopup && (
                <div style={{
                    position: "fixed", top: 0, left: 0, width: "100vw", height: "100vh",
                    background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000
                }}>
                    <div style={{
                        background: "#fff", borderRadius: 12, padding: 32, minWidth: 320, maxWidth: 500, boxShadow: "0 8px 32px rgba(0,0,0,0.15)", textAlign: "center"
                    }}>
                        <h3 style={{ marginBottom: 24 }}>{showProgramPopup}</h3>
                        {(() => {
                            const program = showProgramPopup;
                            const taskOrder = selectedTaskOrders[program] || null;
                            return (
                                <>
                                    {taskOrder && (
                                        <div style={{ fontWeight: 500, color: "#764ba2", marginBottom: 12 }}>
                                            Task Order: <span style={{ fontWeight: 700 }}>{taskOrder}</span>
                                        </div>
                                    )}
                                    {results[program] ? (
                                        <>
                                            <div
                                                className={styles.chatbotScore}
                                                style={{ color: getScoreColor(results[program]?.score || 1), marginBottom: 12 }}
                                            >
                                                Score: {results[program]?.score}/5
                                            </div>
                                            <div className={styles.reasonBox} style={{ marginBottom: 16 }}>
                                                <span className={styles.reasonLabel}>Reason:</span> {results[program]?.reason}
                                            </div>
                                        </>
                                    ) : (
                                        <div style={{ marginBottom: 16, color: "#aaa" }}>No result yet.</div>
                                    )}
                                </>
                            );
                        })()}
                        <button
                            className={styles.startButton}
                            style={{ background: "#6b7280" }}
                            onClick={() => setShowProgramPopup(null)}
                        >
                            Close
                        </button>
                    </div>
                </div>
            )}

            <ul className={styles.programList}>
                {(started ? filteredProgramList : programList).map((program, idx) => {
                    const checked = selectedPrograms[program] ?? true;
                    const taskOrders = hierarchicalPrograms[program] || [];
                    const selectedTaskOrder = selectedTaskOrders[program] || null;
                    const isScoring = started && idx === currentIdx && results[program] == null;

                    return (
                        <li
                            className={styles.programItem}
                            key={program}
                            style={{ cursor: started && results[program] ? 'pointer' : !started ? 'pointer' : undefined }}
                            onClick={e => {
                                // If not started and not clicking on a checkbox, toggle check/uncheck
                                if (!started && (e.target as HTMLElement).tagName !== "INPUT") {
                                    handleProgramCheckbox(program, !checked);
                                }
                                // If started and scoring is complete, show popup (keep your existing logic)
                                if (
                                    started &&
                                    currentIdx >= filteredProgramList.length &&
                                    (e.target as HTMLElement).tagName !== "INPUT"
                                ) {
                                    setShowProgramPopup(program);
                                }
                            }}
                        >
                            {!started && (
                                <input
                                    type="checkbox"
                                    checked={checked}
                                    onChange={e => handleProgramCheckbox(program, e.target.checked)}
                                    style={{ marginRight: 12 }}
                                    onClick={e => e.stopPropagation()}
                                />
                            )}
                            <span className={styles.programName}>{program}</span>
                            {/* Task Orders BELOW the program name */}
                            {taskOrders.length > 0 && !started && (
                                <div style={{ marginLeft: 0, marginTop: 8 }}>
                                    <div style={{ fontWeight: 500, color: "#764ba2", marginBottom: 4 }}>Task Orders:</div>
                                    <ul style={{ margin: 0, paddingLeft: 20 }}>
                                        {taskOrders.map((task, tIdx) => (
                                            <li
                                                key={task + tIdx}
                                                style={{ marginBottom: 4, display: "flex", alignItems: "center", cursor: !checked ? "not-allowed" : "pointer" }}
                                                onClick={e => {
                                                    if (!checked) return;
                                                    setSelectedTaskOrders(prev => ({
                                                        ...prev,
                                                        [program]: prev[program] === task ? null : task
                                                    }));
                                                }}
                                            >
                                                <input
                                                    type="checkbox"
                                                    checked={selectedTaskOrder === task}
                                                    disabled={!checked}
                                                    onChange={() => {
                                                        setSelectedTaskOrders(prev => ({
                                                            ...prev,
                                                            [program]: prev[program] === task ? null : task
                                                        }));
                                                    }}
                                                    style={{ marginRight: 8 }}
                                                    onClick={e => e.stopPropagation()}
                                                />
                                                <span style={{ fontStyle: "italic", color: checked ? undefined : "#aaa" }}>{task}</span>
                                            </li>
                                        ))}
                                    </ul>
                                </div>
                            )}
                            {/* After scoring starts, just show the selected task order (if any) BELOW the program */}
                            {taskOrders.length > 0 && started && (
                                <div style={{ marginLeft: 0, marginTop: 8 }}>
                                    <span style={{ fontWeight: 500, color: "#764ba2" }}>
                                        Task Order:{" "}
                                        <span style={{ fontWeight: 700 }}>
                                            {selectedTaskOrder || <span style={{ color: "#aaa" }}>None selected</span>}
                                        </span>
                                    </span>
                                </div>
                            )}
                            {/* Chatbot answer box is always below program and task order */}
                            {started && (
                                <div className={styles.resultContent} style={{ marginTop: 12 }}>
                                    {results[program] != null ? (
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
                                        <span className={`${styles.status} ${isScoring ? styles.statusActive : ''}`}>
                                            {isScoring ? (
                                                <span>
                                                    <span className={styles.loadingDot}>●</span> Scoring...
                                                </span>
                                            ) : (
                                                "Waiting..."
                                            )}
                                        </span>
                                    )}
                                </div>
                            )}
                        </li>
                    );
                })}
            </ul>
            
            {started && currentIdx >= filteredProgramList.length && (
                <div className="done">
                    ✅ All programs scored for topic: <strong>"{topic}"</strong>
                </div>
            )}

            {showIndexSwitcher && (
    <div style={{
        position: "fixed", top: 0, left: 0, width: "100vw", height: "100vh",
        background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000
    }}>
        <div style={{
            background: "#fff", borderRadius: 12, padding: 32, minWidth: 320, boxShadow: "0 8px 32px rgba(0,0,0,0.15)", textAlign: "center"
        }}>
            <h3 style={{ marginBottom: 24 }}>Switch Microsoft Index</h3>
            <select
                value={pendingIndex}
                onChange={e => setPendingIndex(e.target.value)}
                style={{ padding: 10, borderRadius: 8, minWidth: 220, fontSize: 16, marginBottom: 24 }}
            >
                {indexes.map(idx => (
                    <option key={idx} value={idx}>{idx}</option>
                ))}
            </select>
            <div style={{ marginTop: 16 }}>
                <button
                    className={styles.startButton}
                    style={{ marginRight: 12 }}
                    onClick={() => {
                        setSelectedIndex(pendingIndex);
                        setShowIndexSwitcher(false);
                        setResults({});
                        setCurrentIdx(0);
                        setStarted(false);
                        setProcessing(false);
                    }}
                >
                    Confirm
                </button>
                <button
                    className={styles.startButton}
                    style={{ background: "#6b7280" }}
                    onClick={() => setShowIndexSwitcher(false)}
                >
                    Cancel
                </button>
            </div>
        </div>
    </div>
)}
        </div>
    );
}