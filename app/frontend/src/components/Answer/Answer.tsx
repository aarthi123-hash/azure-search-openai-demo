import { useMemo, useState } from "react";
import { Stack, IconButton } from "@fluentui/react";
import { useTranslation } from "react-i18next";
import DOMPurify from "dompurify";
import { saveAs } from "file-saver";
import jsPDF from "jspdf";
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } from "docx";

import styles from "./Answer.module.css";
import { ChatAppResponse, getCitationFilePath, SpeechConfig } from "../../api";
import { parseAnswerToHtml } from "./AnswerParser";
import { AnswerIcon } from "./AnswerIcon";

interface Props {
    answer: ChatAppResponse;
    index: number;
    speechConfig: SpeechConfig;
    isSelected?: boolean;
    isStreaming: boolean;
    onCitationClicked: (filePath: string) => void;
    onThoughtProcessClicked: () => void;
    onSupportingContentClicked: () => void;
    onFollowupQuestionClicked?: (question: string) => void;
    showFollowupQuestions?: boolean;
    showSpeechOutputBrowser?: boolean;
    showSpeechOutputAzure?: boolean;
}

export const Answer = ({
    answer,
    index,
    speechConfig,
    isSelected,
    isStreaming,
    onCitationClicked,
    onThoughtProcessClicked,
    onSupportingContentClicked,
    onFollowupQuestionClicked,
    showFollowupQuestions,
    showSpeechOutputAzure,
    showSpeechOutputBrowser
}: Props) => {
    const followupQuestions = answer.context?.followup_questions;
    const parsedAnswer = useMemo(() => parseAnswerToHtml(answer, isStreaming, onCitationClicked), [answer]);
    const { t } = useTranslation();

    // Convert markdown-like content to HTML and sanitize it
    const convertToHtml = (content: string) => {
        // Basic markdown to HTML conversion
        let html = content
            // Headers
            .replace(/^### (.*$)/gim, '<h3>$1</h3>')
            .replace(/^## (.*$)/gim, '<h2>$1</h2>')
            .replace(/^# (.*$)/gim, '<h1>$1</h1>')
            // Bold
            .replace(/\*\*(.*)\*\*/gim, '<strong>$1</strong>')
            // Italic
            .replace(/\*(.*)\*/gim, '<em>$1</em>')
            // Code blocks
            .replace(/```([^`]+)```/gim, '<pre><code>$1</code></pre>')
            // Inline code
            .replace(/`([^`]+)`/gim, '<code>$1</code>')
            // Links
            .replace(/\[([^\]]+)\]\(([^)]+)\)/gim, '<a href="$2" target="_blank" rel="noopener noreferrer">$1</a>')
            // Line breaks
            .replace(/\n\n/gim, '</p><p>')
            .replace(/\n/gim, '<br>')
            // Lists
            .replace(/^\* (.+)$/gim, '<li>$1</li>')
            .replace(/(<li>.*<\/li>)/gims, '<ul>$1</ul>')
            .replace(/^\d+\. (.+)$/gim, '<li>$1</li>')
            .replace(/(<li>.*<\/li>)/gims, '<ol>$1</ol>');
        
        // Wrap content in paragraphs if not already wrapped
        if (!html.startsWith('<')) {
            html = `<p>${html}</p>`;
        }
        
        return html;
    };
    
    const htmlContent = convertToHtml(parsedAnswer.answerHtml);

    // Sanitize the HTML content before rendering or copying
    const sanitizedAnswerHtml = DOMPurify.sanitize(htmlContent, {
        ALLOWED_TAGS: [
            'p', 'br', 'strong', 'em', 'u', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
            'ul', 'ol', 'li', 'blockquote', 'code', 'pre', 'a', 'table', 'thead',
            'tbody', 'tr', 'th', 'td', 'div', 'span', 'sup'
        ],
        ALLOWED_ATTR: ['href', 'target', 'rel', 'class', 'id']
    });

    const [copied, setCopied] = useState(false);

    const handleCopy = () => {
        // Remove HTML tags for plain text copy
        const textToCopy = sanitizedAnswerHtml.replace(/<a [^>]*><sup>\d+<\/sup><\/a>|<[^>]+>/g, "");
        
        navigator.clipboard
            .writeText(textToCopy)
            .then(() => {
                setCopied(true);
                setTimeout(() => setCopied(false), 2000);
            })
            .catch(err => console.error("Failed to copy text: ", err));
    };


    const handleDownloadHtml = () => {
        const answerHtml = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Chatbot Answer</title>
</head>
<body>
    ${sanitizedAnswerHtml}
</body>
</html>
        `.trim();

        const blob = new Blob([answerHtml], { type: "text/html;charset=utf-8" });
        saveAs(blob, "chatbot-answer.html");
    };

    const handleDownloadPdf = () => {
        const doc = new jsPDF();
        let yPosition = 20;
        const pageHeight = doc.internal.pageSize.height;
        const margin = 20;
        const maxWidth = 170; // Page width minus margins

        // Helper function to add a new page if needed
        const checkPageBreak = (requiredHeight = 10) => {
            if (yPosition + requiredHeight > pageHeight - margin) {
                doc.addPage();
                yPosition = 20;
            }
        };

        // Helper function to parse HTML and add formatted content
        const addFormattedContent = (htmlContent: string) => {
            // Create a temporary div to parse HTML
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = htmlContent;

            // Process each element
            const processElement = (element: Element) => {
                const tagName = element.tagName?.toLowerCase();
                const textContent = element.textContent || '';

                switch (tagName) {
                    case 'h1':
                        checkPageBreak(15);
                        doc.setFontSize(18);
                        doc.setFont("helvetica", "bold");
                        const h1Lines = doc.splitTextToSize(textContent, maxWidth);
                        doc.text(h1Lines, margin, yPosition);
                        yPosition += h1Lines.length * 8 + 10;
                        break;
                    case 'h2':
                        checkPageBreak(12);
                        doc.setFontSize(16);
                        doc.setFont("helvetica", "bold");
                        const h2Lines = doc.splitTextToSize(textContent, maxWidth);
                        doc.text(h2Lines, margin, yPosition);
                        yPosition += h2Lines.length * 7 + 8;
                        break;
                    case 'h3':
                        checkPageBreak(10);
                        doc.setFontSize(14);
                        doc.setFont("helvetica", "bold");
                        const h3Lines = doc.splitTextToSize(textContent, maxWidth);
                        doc.text(h3Lines, margin, yPosition);
                        yPosition += h3Lines.length * 6 + 6;
                        break;
                    case 'p':
                        checkPageBreak(8);
                        doc.setFontSize(11);
                        doc.setFont("helvetica", "normal");
                        const pLines = doc.splitTextToSize(textContent, maxWidth);
                        doc.text(pLines, margin, yPosition);
                        yPosition += pLines.length * 5 + 5;
                        break;
                    case 'strong':
                    case 'b':
                        doc.setFont("helvetica", "bold");
                        const boldLines = doc.splitTextToSize(textContent, maxWidth);
                        doc.text(boldLines, margin, yPosition);
                        yPosition += boldLines.length * 5;
                        doc.setFont("helvetica", "normal");
                        break;
                    case 'ul':
                        checkPageBreak(5);
                        yPosition += 3;
                        Array.from(element.children).forEach((li, index) => {
                            if (li.tagName.toLowerCase() === 'li') {
                                checkPageBreak(6);
                                doc.setFontSize(11);
                                doc.setFont("helvetica", "normal");
                                const bullet = '• ';
                                const liText = li.textContent || '';
                                const liLines = doc.splitTextToSize(bullet + liText, maxWidth - 10);
                                doc.text(liLines, margin + 5, yPosition);
                                yPosition += liLines.length * 5 + 2;
                            }
                        });
                        yPosition += 3;
                        break;
                    case 'ol':
                        checkPageBreak(5);
                        yPosition += 3;
                        Array.from(element.children).forEach((li, index) => {
                            if (li.tagName.toLowerCase() === 'li') {
                                checkPageBreak(6);
                                doc.setFontSize(11);
                                doc.setFont("helvetica", "normal");
                                const number = `${index + 1}. `;
                                const liText = li.textContent || '';
                                const liLines = doc.splitTextToSize(number + liText, maxWidth - 10);
                                doc.text(liLines, margin + 5, yPosition);
                                yPosition += liLines.length * 5 + 2;
                            }
                        });
                        yPosition += 3;
                        break;
                    case 'code':
                        checkPageBreak(8);
                        doc.setFontSize(10);
                        doc.setFont('courier', 'normal');
                        const codeLines = doc.splitTextToSize(textContent, maxWidth - 10);
                        const codeHeight = codeLines.length * 4 + 4;
                        doc.setFillColor(240, 240, 240);
                        doc.rect(margin, yPosition - 2, maxWidth, codeHeight, 'F');
                        doc.text(codeLines, margin + 2, yPosition + 2);
                        yPosition += codeHeight + 3;
                        doc.setFont("helvetica", 'normal');
                        break;
                    case 'pre':
                        checkPageBreak(10);
                        doc.setFontSize(9);
                        doc.setFont('courier', 'normal');
                        const preLines = doc.splitTextToSize(textContent, maxWidth - 10);
                        const preHeight = preLines.length * 4 + 6;
                        doc.setFillColor(245, 245, 245);
                        doc.rect(margin, yPosition - 2, maxWidth, preHeight, 'F');
                        doc.text(preLines, margin + 3, yPosition + 2);
                        yPosition += preHeight + 5;
                        doc.setFont("helvetica", 'normal');
                        break;
                    case 'blockquote':
                        checkPageBreak(8);
                        doc.setFontSize(11);
                        doc.setFont("helvetica", "italic");
                        doc.setLineWidth(2);
                        doc.setDrawColor(200, 200, 200);
                        const quoteLines = doc.splitTextToSize(textContent, maxWidth - 15);
                        const quoteHeight = quoteLines.length * 5;
                        doc.line(margin, yPosition - 2, margin, yPosition + quoteHeight);
                        doc.text(quoteLines, margin + 8, yPosition);
                        yPosition += quoteHeight + 5;
                        doc.setFont("helvetica", 'normal');
                        break;
                    default:
                        if (textContent.trim()) {
                            checkPageBreak(6);
                            doc.setFontSize(11);
                            doc.setFont("helvetica", "normal");
                            const defaultLines = doc.splitTextToSize(textContent, maxWidth);
                            doc.text(defaultLines, margin, yPosition);
                            yPosition += defaultLines.length * 5 + 3;
                        }
                        break;
                }
            };

            Array.from(tempDiv.children).forEach(processElement);
        };

        // Add title
        doc.setFontSize(20);
        doc.setFont("helvetica", "bold");
        doc.text('Chatbot Response', margin, yPosition);
        yPosition += 15;

    

        // Add a separator line
        doc.setLineWidth(0.5);
        doc.setDrawColor(100, 100, 100);
        doc.line(margin, yPosition, maxWidth + margin, yPosition);
        yPosition += 10;

        // Add the formatted content
        addFormattedContent(sanitizedAnswerHtml);

        // Add footer with page numbers
        //const pageCount = doc.getNumberOfPages();
       // for (let i = 1; i <= pageCount; i++) {
        //    doc.setPage(i);
        //    doc.setFontSize(10);
            doc.setFont("helvetica", "normal");
      //      doc.text(`Page ${i} of ${pageCount}`, maxWidth + margin - 30, pageHeight - 10);
       //     doc.text(`Generated on ${new Date().toLocaleDateString()}`, margin, pageHeight - 10);
     //   }
        doc.save("chatbot-answer.pdf");
    };

    const handleDownloadWord = async () => {
        // Helper to parse HTML and create docx Paragraphs
        const htmlToDocxParagraphs = (htmlContent: string) => {
            const tempDiv = document.createElement("div");
            tempDiv.innerHTML = htmlContent;
            const paragraphs: Paragraph[] = [];

            const processElement = (element: Element) => {
                const tagName = element.tagName?.toLowerCase();
                const textContent = element.textContent || "";

                switch (tagName) {
                    case "h1":
                        paragraphs.push(
                            new Paragraph({
                                text: textContent,
                                heading: HeadingLevel.HEADING_1,
                                spacing: { after: 200 }
                            })
                        );
                        break;
                    case "h2":
                        paragraphs.push(
                            new Paragraph({
                                text: textContent,
                                heading: HeadingLevel.HEADING_2,
                                spacing: { after: 150 }
                            })
                        );
                        break;
                    case "h3":
                        paragraphs.push(
                            new Paragraph({
                                text: textContent,
                                heading: HeadingLevel.HEADING_3,
                                spacing: { after: 100 }
                            })
                        );
                        break;
                    case "p":
                        paragraphs.push(
                            new Paragraph({
                                children: [new TextRun({ text: textContent, font: "Arial", size: 22 })],
                                spacing: { after: 100 }
                            })
                        );
                        break;
                    case "strong":
                    case "b":
                        paragraphs.push(
                            new Paragraph({
                                children: [new TextRun({ text: textContent, bold: true, font: "Arial", size: 22 })]
                            })
                        );
                        break;
                    case "ul":
                        Array.from(element.children).forEach((li) => {
                            if (li.tagName.toLowerCase() === "li") {
                                paragraphs.push(
                                    new Paragraph({
                                        text: (li.textContent || ""),
                                        bullet: { level: 0 }
                                    })
                            );
                            }
                        });
                        break;
                    case "ol":
                        Array.from(element.children).forEach((li, idx) => {
                            if (li.tagName.toLowerCase() === "li") {
                                paragraphs.push(
                                    new Paragraph({
                                        text: `${idx + 1}. ${(li.textContent || "")}`,
                                        numbering: { reference: "numbered-list", level: 0 }
                                    })
                            );
                            }
                        });
                        break;
                    case "blockquote":
                        paragraphs.push(
                            new Paragraph({
                                children: [
                                    new TextRun({ text: textContent, italics: true })
                                ],
                                alignment: AlignmentType.LEFT,
                                border: { 
                                    left: { 
                                        color: "auto", 
                                        space: 1, 
                                        size: 6, 
                                        style: "single"
                                    } 
                                },
                                spacing: { after: 100 }
                            })
                        );
                        break;
                    case "code":
                    case "pre":
                        paragraphs.push(
                            new Paragraph({
                                children: [new TextRun({ text: textContent, font: "Courier New", size: 20 })],
                                shading: { fill: "EDEDED" },
                                spacing: { after: 100 }
                            })
                        );
                        break;
                    default:
                        if (textContent.trim()) {
                            paragraphs.push(
                                new Paragraph({
                                    children: [new TextRun({ text: textContent, font: "Arial", size: 22 })],
                                    spacing: { after: 100 }
                            })
                        );
                        }
                        break;
                }
            };

            Array.from(tempDiv.children).forEach(processElement);
            return paragraphs;
        };

        // Build the document
        const doc = new Document({
            sections: [
                {
                    properties: {},
                    children: [
                        new Paragraph({
                            text: "Chatbot Response",
                            heading: HeadingLevel.TITLE,
                            alignment: AlignmentType.LEFT,
                            spacing: { after: 200 }
                        }),
                        ...htmlToDocxParagraphs(sanitizedAnswerHtml),
                        new Paragraph({
                            text: `Generated on ${new Date().toLocaleDateString()}`,
                            alignment: AlignmentType.LEFT,
                            spacing: { before: 400 }
                        })
                    ]
                }
            ]
        });

        const blob = await Packer.toBlob(doc);
        saveAs(blob, "chatbot-answer.docx");
    };

    return (
        <div>
            <Stack className={`${styles.answerContainer} ${isSelected && styles.selected}`} verticalAlign="space-between">
                <Stack.Item>
                    <Stack horizontal horizontalAlign="space-between">
                        <AnswerIcon />
                        <div>
                            <IconButton
                                style={{ color: "black" }}
                                iconProps={{ iconName: copied ? "CheckMark" : "Copy" }}
                                title={copied ? t("tooltips.copied") : t("tooltips.copy")}
                                ariaLabel={copied ? t("tooltips.copied") : t("tooltips.copy")}
                                onClick={handleCopy}
                            />
                            <IconButton
                                style={{ color: "black" }}
                                iconProps={{ iconName: "Lightbulb" }}
                                title={t("tooltips.showThoughtProcess")}
                                ariaLabel={t("tooltips.showThoughtProcess")}
                                onClick={() => onThoughtProcessClicked()}
                                disabled={!answer.context.thoughts?.length || isStreaming}
                            />
                            <IconButton
                                style={{ color: "black" }}
                                iconProps={{ iconName: "ClipboardList" }}
                                title={t("tooltips.showSupportingContent")}
                                ariaLabel={t("tooltips.showSupportingContent")}
                                onClick={() => onSupportingContentClicked()}
                                disabled={!answer.context.data_points || isStreaming}
                            />
                            <IconButton
                                style={{ color: "black" }}
                                iconProps={{ iconName: "Download" }}
                                title="Download as HTML"
                                ariaLabel="Download as HTML"
                                onClick={handleDownloadHtml}
                            />
                            <IconButton
                                style={{ color: "black" }}
                                iconProps={{ iconName: "PDF" }}
                                title="Download as PDF"
                                ariaLabel="Download as PDF"
                                onClick={handleDownloadPdf}
                            />
                            <IconButton
                                style={{ color: "black" }}
                                iconProps={{ iconName: "WordLogo" }}
                                title="Download as Word"
                                ariaLabel="Download as Word"
                                onClick={handleDownloadWord}
                            />
                        </div>
                    </Stack>
                </Stack.Item>

                <Stack.Item grow>
                    <div 
                        className={styles.answerText}
                        dangerouslySetInnerHTML={{ __html: sanitizedAnswerHtml }}
                    />
                </Stack.Item>

                {!!parsedAnswer.citations.length && (
                    <Stack.Item>
                        <Stack horizontal wrap tokens={{ childrenGap: 5 }}>
                            <span className={styles.citationLearnMore}>{t("citationWithColon")}</span>
                            {parsedAnswer.citations.map((x, i) => {
                                const path = getCitationFilePath(x);
                                return (
                                    <a key={i} className={styles.citation} title={x} onClick={() => onCitationClicked(path)}>
                                        {`${++i}. ${x}`}
                                    </a>
                                );
                            })}
                        </Stack>
                    </Stack.Item>
                )}

                {!!followupQuestions?.length && showFollowupQuestions && onFollowupQuestionClicked && (
                    <Stack.Item>
                        <Stack horizontal wrap className={`${!!parsedAnswer.citations.length ? styles.followupQuestionsList : ""}`} tokens={{ childrenGap: 6 }}>
                            <span className={styles.followupQuestionLearnMore}>{t("followupQuestions")}</span>
                            {followupQuestions.map((x, i) => {
                                return (
                                    <a key={i} className={styles.followupQuestion} title={x} onClick={() => onFollowupQuestionClicked(x)}>
                                        {`${x}`}
                                    </a>
                                );
                            })}
                        </Stack>
                    </Stack.Item>
                )}
            </Stack>
        </div>
    )
}