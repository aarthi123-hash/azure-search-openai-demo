import { useContext, useEffect, useRef, useState } from "react";
import { useTranslation } from "react-i18next";
import { Helmet } from "react-helmet-async";
import { 
    Panel, 
    DefaultButton, 
    Spinner, 
    Dropdown, 
    IDropdownOption, 
    Stack, 
    IStackTokens, 
    IconButton, 
    Text,
    MessageBar,
    MessageBarType,
    Separator,
    Icon
} from "@fluentui/react";
import mammoth from "mammoth"; // If you want to support .docx parsing

import styles from "./Ask.module.css";

import { askApi, configApi, uploadFileApi, ChatAppResponse, ChatAppRequest, RetrievalMode, VectorFields, GPT4VInput, SpeechConfig } from "../../api";
import { Answer, AnswerError } from "../../components/Answer";
import { QuestionInput } from "../../components/QuestionInput";
import { ExampleList } from "../../components/Example";
import { AnalysisPanel, AnalysisPanelTabs } from "../../components/AnalysisPanel";
import { SettingsButton } from "../../components/SettingsButton/SettingsButton";
import { useLogin, getToken, requireAccessControl } from "../../authConfig";
import { UploadFile } from "../../components/UploadFile";
import { Settings } from "../../components/Settings/Settings";
import { useMsal } from "@azure/msal-react";
import { TokenClaimsDisplay } from "../../components/TokenClaimsDisplay";
import { LoginContext } from "../../loginContext";
import { LanguagePicker } from "../../i18n/LanguagePicker";
import { getProgramOptions, getTaskOrderOptions, hasTaskOrders } from "./filterUtils";
import { filterConfig, FilterConfig } from "./filterConfig";

// Enhanced styles for the filter components
const filterStyles = {
    filterContainer: {
        background: 'linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)',
        padding: '20px',
        borderRadius: '12px',
        marginBottom: '20px',
        border: '2px solid #e1e8ed',
        boxShadow: '0 4px 12px rgba(0, 0, 0, 0.1)'
    },
    filterHeader: {
        marginBottom: '16px',
        display: 'flex',
        alignItems: 'center',
        gap: '8px'
    },
    filterTitle: {
        fontSize: '16px',
        fontWeight: '600',
        color: '#323130',
        margin: 0
    },
    filterDropdown: {
        minWidth: '220px'
    },
    filterTagsContainer: {
        marginTop: '16px',
        padding: '12px',
        backgroundColor: '#ffffff',
        borderRadius: '8px',
        border: '1px solid #e1e8ed'
    },
    filterTag: {
        display: 'inline-flex',
        alignItems: 'center',
        background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
        color: '#ffffff',
        padding: '6px 12px',
        borderRadius: '20px',
        fontSize: '13px',
        fontWeight: '500',
        marginRight: '8px',
        marginBottom: '8px',
        boxShadow: '0 2px 4px rgba(0, 0, 0, 0.1)',
        transition: 'all 0.2s ease'
    },
    removeButton: {
        marginLeft: '6px',
        minWidth: '18px',
        height: '18px',
        color: '#ffffff',
        backgroundColor: 'rgba(255, 255, 255, 0.2)',
        borderRadius: '50%',
        border: 'none',
        cursor: 'pointer',
        ':hover': {
            backgroundColor: 'rgba(255, 255, 255, 0.3)'
        }
    },
    activeFiltersHeader: {
        display: 'flex',
        alignItems: 'center',
        gap: '8px',
        marginBottom: '8px'
    },
    activeFiltersTitle: {
        fontSize: '14px',
        fontWeight: '600',
        color: '#605e5c'
    },
    filterSummary: {
        backgroundColor: '#f3f9ff',
        border: '1px solid #cfe4fd',
        borderRadius: '8px',
        padding: '12px',
        marginBottom: '16px'
    },
    filterSummaryTitle: {
        fontSize: '14px',
        fontWeight: '600',
        color: '#0078d4',
        marginBottom: '8px',
        display: 'flex',
        alignItems: 'center',
        gap: '6px'
    },
    filterSummaryContent: {
        fontSize: '13px',
        color: '#323130',
        lineHeight: '1.4'
    },
    clearAllButton: {
        background: 'linear-gradient(135deg, #ff6b6b 0%, #ee5a52 100%)',
        color: '#ffffff',
        border: 'none',
        borderRadius: '6px',
        padding: '8px 16px',
        fontWeight: '500',
        cursor: 'pointer',
        transition: 'all 0.2s ease',
        ':hover': {
            transform: 'translateY(-1px)',
            boxShadow: '0 4px 8px rgba(238, 90, 82, 0.3)'
        }
    }
};

const stackTokens: IStackTokens = { childrenGap: 20 };

interface ActiveFilters {
    program?: string;
    taskOrder?: string;
}

type ChatAppResponseWithFilter = ChatAppResponse & { filterSummary?: string; session_state?: any };

export function Component(): JSX.Element {
    // ... (all your existing state variables remain the same)
    const [isConfigPanelOpen, setIsConfigPanelOpen] = useState(false);
    const [promptTemplate, setPromptTemplate] = useState<string>("");
    const [promptTemplatePrefix, setPromptTemplatePrefix] = useState<string>("");
    const [promptTemplateSuffix, setPromptTemplateSuffix] = useState<string>("");
    const [temperature, setTemperature] = useState<number>(0.3);
    const [seed, setSeed] = useState<number | null>(null);
    const [minimumRerankerScore, setMinimumRerankerScore] = useState<number>(0);
    const [minimumSearchScore, setMinimumSearchScore] = useState<number>(0);
    const [retrievalMode, setRetrievalMode] = useState<RetrievalMode>(RetrievalMode.Hybrid);
    const [retrieveCount, setRetrieveCount] = useState<number>(3);
    const [maxSubqueryCount, setMaxSubqueryCount] = useState<number>(10);
    const [resultsMergeStrategy, setResultsMergeStrategy] = useState<string>("interleaved");
    const [useSemanticRanker, setUseSemanticRanker] = useState<boolean>(true);
    const [useSemanticCaptions, setUseSemanticCaptions] = useState<boolean>(false);
    const [useQueryRewriting, setUseQueryRewriting] = useState<boolean>(false);
    const [reasoningEffort, setReasoningEffort] = useState<string>("");
    const [useGPT4V, setUseGPT4V] = useState<boolean>(false);
    const [gpt4vInput, setGPT4VInput] = useState<GPT4VInput>(GPT4VInput.TextAndImages);
    const [includeCategory, setIncludeCategory] = useState<string>("");
    const [excludeCategory, setExcludeCategory] = useState<string>("");
    const [question, setQuestion] = useState<string>("");
    const [vectorFields, setVectorFields] = useState<VectorFields>(VectorFields.TextAndImageEmbeddings);
    const [useOidSecurityFilter, setUseOidSecurityFilter] = useState<boolean>(false);
    const [useGroupsSecurityFilter, setUseGroupsSecurityFilter] = useState<boolean>(false);
    const [showGPT4VOptions, setShowGPT4VOptions] = useState<boolean>(false);
    const [showSemanticRankerOption, setShowSemanticRankerOption] = useState<boolean>(false);
    const [showQueryRewritingOption, setShowQueryRewritingOption] = useState<boolean>(false);
    const [showReasoningEffortOption, setShowReasoningEffortOption] = useState<boolean>(false);
    const [showVectorOption, setShowVectorOption] = useState<boolean>(false);
    const [showUserUpload, setShowUserUpload] = useState<boolean>(false);
    const [showLanguagePicker, setshowLanguagePicker] = useState<boolean>(false);
    const [showSpeechInput, setShowSpeechInput] = useState<boolean>(false);
    const [showSpeechOutputBrowser, setShowSpeechOutputBrowser] = useState<boolean>(false);
    const [showSpeechOutputAzure, setShowSpeechOutputAzure] = useState<boolean>(false);
    const audio = useRef(new Audio()).current;
    const [isPlaying, setIsPlaying] = useState(false);
    const [showAgenticRetrievalOption, setShowAgenticRetrievalOption] = useState<boolean>(false);
    const [useAgenticRetrieval, setUseAgenticRetrieval] = useState<boolean>(false);

    // Filter states
    const [activeFilters, setActiveFilters] = useState<ActiveFilters>({});
    const [selectedProgram, setSelectedProgram] = useState<string>("");
    const [selectedTaskOrder, setSelectedTaskOrder] = useState<string>("");

    const lastQuestionRef = useRef<string>("");
    const [isLoading, setIsLoading] = useState<boolean>(false);
    const [error, setError] = useState<unknown>();
    const [answer, setAnswer] = useState<ChatAppResponseWithFilter>();
    const [speechUrls, setSpeechUrls] = useState<(string | null)[]>([]);

    const speechConfig: SpeechConfig = {
        speechUrls,
        setSpeechUrls,
        audio,
        isPlaying,
        setIsPlaying
    };

    const [activeCitation, setActiveCitation] = useState<string>();
    const [activeAnalysisPanelTab, setActiveAnalysisPanelTab] = useState<AnalysisPanelTabs | undefined>(undefined);

    const client = useLogin ? useMsal().instance : undefined;
    const { loggedIn } = useContext(LoginContext);

    const getConfig = async () => {
        configApi().then(config => {
            setShowGPT4VOptions(config.showGPT4VOptions);
            setUseSemanticRanker(config.showSemanticRankerOption);
            setShowSemanticRankerOption(config.showSemanticRankerOption);
            setUseQueryRewriting(config.showQueryRewritingOption);
            setShowQueryRewritingOption(config.showQueryRewritingOption);
            setShowReasoningEffortOption(config.showReasoningEffortOption);
            if (config.showReasoningEffortOption) {
                setReasoningEffort(config.defaultReasoningEffort);
            }
            setShowVectorOption(config.showVectorOption);
            if (!config.showVectorOption) {
                setRetrievalMode(RetrievalMode.Text);
            }
            setShowUserUpload(config.showUserUpload);
            setshowLanguagePicker(config.showLanguagePicker);
            setShowSpeechInput(config.showSpeechInput);
            setShowSpeechOutputBrowser(config.showSpeechOutputBrowser);
            setShowSpeechOutputAzure(config.showSpeechOutputAzure);
            setShowAgenticRetrievalOption(config.showAgenticRetrievalOption);
            setUseAgenticRetrieval(config.showAgenticRetrievalOption);
            if (config.showAgenticRetrievalOption) {
                setRetrieveCount(10);
            }
        });
    };

    useEffect(() => {
        getConfig();
    }, []);

    // Handle program dropdown change
    const handleProgramChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
        if (option) {
            const programKey = option.key as string;
            setSelectedProgram(programKey);
            
            // Clear task order if the new program doesn't have task orders
            if (!hasTaskOrders(programKey, currentConfig ?? undefined)) {
                setSelectedTaskOrder("");
                setActiveFilters({ program: programKey });
            } else {
                setActiveFilters({ program: programKey, taskOrder: selectedTaskOrder });
            }
        }
    };

    // Handle task order dropdown change
    const handleTaskOrderChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
        if (option) {
            const taskOrderKey = option.key as string;
            setSelectedTaskOrder(taskOrderKey);
            setActiveFilters({ program: selectedProgram, taskOrder: taskOrderKey });
        }
    };

    // Remove individual filter
    const removeFilter = (filterType: 'program' | 'taskOrder') => {
        if (filterType === 'program') {
            setSelectedProgram("");
            setSelectedTaskOrder("");
            setActiveFilters({});
        } else if (filterType === 'taskOrder') {
            setSelectedTaskOrder("");
            setActiveFilters({ program: selectedProgram });
        }
    };

    // Clear all filters
    const clearAllFilters = () => {
        setSelectedProgram("");
        setSelectedTaskOrder("");
        setActiveFilters({});
    };

    // Build filter query string
    const buildFilterQuery = () => {
        const filters: string[] = [];
        
        if (activeFilters.program) {
            filters.push(`program:${activeFilters.program}`);
        }
        
        if (activeFilters.taskOrder) {
            filters.push(`task_order:${activeFilters.taskOrder}`);
        }
        
        return filters.length > 0 ? `[${filters.join(', ')}]` : '';
    };

    // Get filter summary for display
    const getFilterSummary = (): string | undefined => {
        if (!activeFilters.program && !activeFilters.taskOrder) {
            return undefined;
        }

        const parts: string[] = [];
        if (activeFilters.program) {
            const programLabel = currentConfig && currentConfig.programs
                ? currentConfig.programs[activeFilters.program]?.label || activeFilters.program
                : activeFilters.program;
            parts.push(`Program: ${programLabel}`);
        }
        if (activeFilters.taskOrder) {
            parts.push(`Task Order: ${activeFilters.taskOrder}`);
        }
        
        return parts.join(' • ');
    };

    const makeApiRequest = async (question: string) => {
        const filterQuery = buildFilterQuery();
        const fullQuery = filterQuery ? `${filterQuery} ${question}` : question;
        
        lastQuestionRef.current = fullQuery;

        error && setError(undefined);
        setIsLoading(true);
        setActiveCitation(undefined);
        setActiveAnalysisPanelTab(undefined);

        const token = client ? await getToken(client) : undefined;

        try {
            const request: ChatAppRequest = {
                messages: [
                    {
                        content: fullQuery,
                        role: "user"
                    }
                ],
                context: {
                    overrides: {
                        prompt_template: promptTemplate.length === 0 ? undefined : promptTemplate,
                        prompt_template_prefix: promptTemplatePrefix.length === 0 ? undefined : promptTemplatePrefix,
                        prompt_template_suffix: promptTemplateSuffix.length === 0 ? undefined : promptTemplateSuffix,
                        include_category: includeCategory.length === 0 ? undefined : includeCategory,
                        exclude_category: excludeCategory.length === 0 ? undefined : excludeCategory,
                        top: retrieveCount,
                        max_subqueries: maxSubqueryCount,
                        results_merge_strategy: resultsMergeStrategy,
                        temperature: temperature,
                        minimum_reranker_score: minimumRerankerScore,
                        minimum_search_score: minimumSearchScore,
                        retrieval_mode: retrievalMode,
                        semantic_ranker: useSemanticRanker,
                        semantic_captions: useSemanticCaptions,
                        query_rewriting: useQueryRewriting,
                        reasoning_effort: reasoningEffort,
                        use_oid_security_filter: useOidSecurityFilter,
                        use_groups_security_filter: useGroupsSecurityFilter,
                        vector_fields: vectorFields,
                        use_gpt4v: useGPT4V,
                        gpt4v_input: gpt4vInput,
                        language: i18n.language,
                        use_agentic_retrieval: useAgenticRetrieval,
                        program_filter: activeFilters.program ?? "",
                        task_order_filter: activeFilters.taskOrder ?? "",
                        ...(seed !== null ? { seed: seed } : {})
                    }
                },
                session_state: answer ? answer.session_state : null
            };
            const result = await askApi(request, token);
            
            // After you get the answer from the API:
            const filterSummary = getFilterSummary();
            setAnswer({
                ...result,
                filterSummary // add this property
            });
            setSpeechUrls([null]);
        } catch (e) {
            setError(e);
        } finally {
            setIsLoading(false);
        }
    };

    const handleSend = async (question: string, file?: File | null) => {
        if (file) {
            const formData = new FormData();
            formData.append("file", file);
            try {
                await uploadFileApi(formData);
            } catch (err) {
                // Handle error
            }
        }
        makeApiRequest(question);
    };

    // ... (all your existing handler functions remain the same)

    const onExampleClicked = (example: string) => {
        makeApiRequest(example);
        setQuestion(example);
    };

    const onShowCitation = (citation: string) => {
        if (activeCitation === citation && activeAnalysisPanelTab === AnalysisPanelTabs.CitationTab) {
            setActiveAnalysisPanelTab(undefined);
        } else {
            setActiveCitation(citation);
            setActiveAnalysisPanelTab(AnalysisPanelTabs.CitationTab);
        }
    };

    const onToggleTab = (tab: AnalysisPanelTabs) => {
        if (activeAnalysisPanelTab === tab) {
            setActiveAnalysisPanelTab(undefined);
        } else {
            setActiveAnalysisPanelTab(tab);
        }
    };

    const onUseOidSecurityFilterChange = (_ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) => {
        setUseOidSecurityFilter(!!checked);
    };

    const onUseGroupsSecurityFilterChange = (_ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) => {
        setUseGroupsSecurityFilter(!!checked);
    };

    const handleSettingsChange = (field: string, value: any) => {
        switch (field) {
            case "promptTemplate":
                setPromptTemplate(value);
                break;
            case "promptTemplatePrefix":
                setPromptTemplatePrefix(value);
                break;
            case "promptTemplateSuffix":
                setPromptTemplateSuffix(value);
                break;
            case "temperature":
                setTemperature(value);
                break;
            case "seed":
                setSeed(value);
                break;
            case "minimumRerankerScore":
                setMinimumRerankerScore(value);
                break;
            case "minimumSearchScore":
                setMinimumSearchScore(value);
                break;
            case "retrieveCount":
                setRetrieveCount(value);
                break;
            case "maxSubqueryCount":
                setMaxSubqueryCount(value);
                break;
            case "resultsMergeStrategy":
                setResultsMergeStrategy(value);
                break;
            case "useSemanticRanker":
                setUseSemanticRanker(value);
                break;
            case "useSemanticCaptions":
                setUseSemanticCaptions(value);
                break;
            case "useQueryRewriting":
                setUseQueryRewriting(value);
                break;
            case "reasoningEffort":
                setReasoningEffort(value);
                break;
            case "excludeCategory":
                setExcludeCategory(value);
                break;
            case "includeCategory":
                setIncludeCategory(value);
                break;
            case "useOidSecurityFilter":
                setUseOidSecurityFilter(value);
                break;
            case "useGroupsSecurityFilter":
                setUseGroupsSecurityFilter(value);
                break;
            case "useGPT4V":
                setUseGPT4V(value);
                break;
            case "gpt4vInput":
                setGPT4VInput(value);
                break;
            case "vectorFields":
                setVectorFields(value);
                break;
            case "retrievalMode":
                setRetrievalMode(value);
                break;
            case "useAgenticRetrieval":
                setUseAgenticRetrieval(value);
        }
    };

    const { t, i18n } = useTranslation();

    // New state for dynamic filter config
    const [dynamicFilterConfig, setDynamicFilterConfig] = useState<FilterConfig | null>(null);
    const currentConfig = dynamicFilterConfig ?? filterConfig;

    // Load config from localStorage on mount
    useEffect(() => {
        const savedConfig = localStorage.getItem("defaultFilterConfig");
        if (savedConfig) {
            try {
                setDynamicFilterConfig(JSON.parse(savedConfig));
            } catch {
                // Ignore parse errors
            }
        }
    }, []);

    // Handle config upload
    const handleConfigUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;

        let configObj: any = null;

        if (file.name.endsWith(".json")) {
            const text = await file.text();
            configObj = JSON.parse(text);
        } else if (file.name.endsWith(".docx")) {
            const arrayBuffer = await file.arrayBuffer();
            const result = await mammoth.extractRawText({ arrayBuffer });
            // New parsing logic: each line is either a program or '- taskOrder' under the previous program
            const lines = result.value.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
            const programs: Record<string, { label: string; taskOrders?: string[] }> = {};
            let currentProgram: string | null = null;
            for (const line of lines) {
                if (line.startsWith('-')) {
                    // Task order line
                    if (currentProgram) {
                        const taskOrder = line.replace(/^-\s*/, '');
                        if (!programs[currentProgram].taskOrders) programs[currentProgram].taskOrders = [];
                        programs[currentProgram].taskOrders!.push(taskOrder);
                    }
                } else {
                    // Program line
                    currentProgram = line;
                    programs[currentProgram] = { label: currentProgram };
                }
            }
            configObj = { programs };
        } else {
            alert("Unsupported file type.");
            return;
        }

        setDynamicFilterConfig(configObj);
        localStorage.setItem("defaultFilterConfig", JSON.stringify(configObj)); // Save as new default
    };

    // If no dynamic config, show upload prompt


    return (
        <>
            <div className={styles.askContainer}>
                <Helmet>
                    <title>{t("pageTitle")}</title>
                </Helmet>
                <div className={styles.askTopSection}>
                    <div className={styles.commandsContainer}>
                        {showUserUpload && <UploadFile className={styles.commandButton} disabled={!loggedIn} />}
                        <SettingsButton className={styles.commandButton} onClick={() => setIsConfigPanelOpen(!isConfigPanelOpen)} />
                    </div>
                    <h1 className={styles.askTitle}>{t("askTitle")}</h1>

                    {/* Enhanced Filter Section */}
                    <div style={filterStyles.filterContainer}>
                        <div style={filterStyles.filterHeader}>
                            <Icon iconName="Filter" style={{ color: '#0078d4', fontSize: '16px' }} />
                            <h3 style={filterStyles.filterTitle}>Search Filters</h3>
                        </div>

                        <Stack horizontal tokens={stackTokens} verticalAlign="end">
                            <Stack.Item>
                                <Dropdown
                                    placeholder="Select Program"
                                    label="Program"
                                    options={getProgramOptions(currentConfig)}
                                    selectedKey={selectedProgram}
                                    onChange={handleProgramChange}
                                    styles={{
                                        dropdown: filterStyles.filterDropdown,
                                        title: { borderRadius: '6px' }
                                    }}
                                />
                            </Stack.Item>

                            {selectedProgram && hasTaskOrders(selectedProgram, currentConfig ?? undefined) && (
                                <Stack.Item>
                                    <Dropdown
                                        placeholder="Select Task Order"
                                        label="Task Order"
                                        options={getTaskOrderOptions(selectedProgram, currentConfig ?? undefined)}
                                        selectedKey={selectedTaskOrder}
                                        onChange={handleTaskOrderChange}
                                        styles={{
                                            dropdown: filterStyles.filterDropdown,
                                            title: { borderRadius: '6px' }
                                        }}
                                    />
                                </Stack.Item>
                            )}

                            {(activeFilters.program || activeFilters.taskOrder) && (
                                <Stack.Item>
                                    <DefaultButton
                                        text="Clear All Filters"
                                        onClick={clearAllFilters}
                                        iconProps={{ iconName: 'ClearFilter' }}
                                        styles={{
                                            root: filterStyles.clearAllButton,
                                            rootHovered: {
                                                ...filterStyles.clearAllButton,
                                                transform: 'translateY(-1px)',
                                                boxShadow: '0 4px 8px rgba(238, 90, 82, 0.3)'
                                            }
                                        }}
                                    />
                                </Stack.Item>
                            )}
                        </Stack>

                        {/* Active Filter Tags */}
                        {(activeFilters.program || activeFilters.taskOrder) && (
                            <div style={filterStyles.filterTagsContainer}>
                                <div style={filterStyles.activeFiltersHeader}>
                                    <Icon iconName="CheckboxComposite" style={{ color: '#0078d4', fontSize: '14px' }} />
                                    <Text style={filterStyles.activeFiltersTitle}>
                                        Active Filters
                                    </Text>
                                </div>
                                <div>
                                    {activeFilters.program && (
                                        <span style={filterStyles.filterTag}>
                                            <Icon iconName="Program" style={{ marginRight: '4px', fontSize: '12px' }} />
                                            Program: {(currentConfig && currentConfig.programs && activeFilters.program && currentConfig.programs[activeFilters.program]?.label) || activeFilters.program}
                                            <button
                                                onClick={() => removeFilter('program')}
                                                style={filterStyles.removeButton}
                                                title="Remove program filter"
                                            >
                                                <Icon iconName="Cancel" style={{ fontSize: '10px' }} />
                                            </button>
                                        </span>
                                    )}
                                    {activeFilters.taskOrder && (
                                        <span style={filterStyles.filterTag}>
                                            <Icon iconName="Task" style={{ marginRight: '4px', fontSize: '12px' }} />
                                            Task Order: {activeFilters.taskOrder}
                                            <button
                                                onClick={() => removeFilter('taskOrder')}
                                                style={filterStyles.removeButton}
                                                title="Remove task order filter"
                                            >
                                                <Icon iconName="Cancel" style={{ fontSize: '10px' }} />
                                            </button>
                                        </span>
                                    )}
                                </div>
                            </div>
                        )}
                    </div>

                    <div className={styles.askQuestionInput}>
                        <QuestionInput
                            clearOnSend
                            placeholder={t("defaultExamples.placeholder")}
                            disabled={isLoading}
                            onSend={handleSend}
                            showSpeechInput={showSpeechInput}
                        />
                    </div>
                </div>

                <div className={styles.askBottomSection}>
                    {isLoading && <Spinner label={t("generatingAnswer")} />}
                    {!lastQuestionRef.current && (
                        <div className={styles.askTopSection}>
                            {showLanguagePicker && <LanguagePicker onLanguageChange={newLang => i18n.changeLanguage(newLang)} />}
                            <ExampleList onExampleClicked={onExampleClicked} useGPT4V={useGPT4V} />
                        </div>
                    )}

                    {/* Enhanced Answer Section with Filter Display */}
                    {!isLoading && answer && !error && (
                        <div className={styles.askAnswerContainer}>
                            {/* Filter Summary Display */}
                            {getFilterSummary() && (
                                <div style={filterStyles.filterSummary}>
                                    <div style={filterStyles.filterSummaryTitle}>
                                        <Icon iconName="FilterSolid" />
                                        Filtered Results
                                    </div>
                                    <div style={filterStyles.filterSummaryContent}>
                                        Results filtered by: {getFilterSummary()}
                                    </div>
                                </div>
                            )}

                            {/* Answer Component */}
                            <Answer
                                answer={answer}
                                index={0}
                                speechConfig={speechConfig}
                                isStreaming={false}
                                onCitationClicked={x => onShowCitation(x)}
                                onThoughtProcessClicked={() => onToggleTab(AnalysisPanelTabs.ThoughtProcessTab)}
                                onSupportingContentClicked={() => onToggleTab(AnalysisPanelTabs.SupportingContentTab)}
                                showSpeechOutputAzure={showSpeechOutputAzure}
                                showSpeechOutputBrowser={showSpeechOutputBrowser}
                            />
                        </div>
                    )}

                    {error ? (
                        <div className={styles.askAnswerContainer}>
                            <AnswerError error={error.toString()} onRetry={() => makeApiRequest(lastQuestionRef.current)} />
                        </div>
                    ) : null}

                    {activeAnalysisPanelTab && answer && (
                        <AnalysisPanel
                            className={styles.askAnalysisPanel}
                            activeCitation={activeCitation}
                            onActiveTabChanged={x => onToggleTab(x)}
                            citationHeight="600px"
                            answer={answer}
                            activeTab={activeAnalysisPanelTab}
                        />
                    )}
                </div>

                <Panel
                    headerText={t("labels.headerText")}
                    isOpen={isConfigPanelOpen}
                    isBlocking={false}
                    onDismiss={() => setIsConfigPanelOpen(false)}
                    closeButtonAriaLabel={t("labels.closeButton")}
                    onRenderFooterContent={() => <DefaultButton onClick={() => setIsConfigPanelOpen(false)}>{t("labels.closeButton")}</DefaultButton>}
                    isFooterAtBottom={true}
                >
                    <Settings
                        promptTemplate={promptTemplate}
                        promptTemplatePrefix={promptTemplatePrefix}
                        promptTemplateSuffix={promptTemplateSuffix}
                        temperature={temperature}
                        retrieveCount={retrieveCount}
                        maxSubqueryCount={maxSubqueryCount}
                        resultsMergeStrategy={resultsMergeStrategy}
                        seed={seed}
                        minimumSearchScore={minimumSearchScore}
                        minimumRerankerScore={minimumRerankerScore}
                        useSemanticRanker={useSemanticRanker}
                        useSemanticCaptions={useSemanticCaptions}
                        useQueryRewriting={useQueryRewriting}
                        reasoningEffort={reasoningEffort}
                        excludeCategory={excludeCategory}
                        includeCategory={includeCategory}
                        retrievalMode={retrievalMode}
                        useGPT4V={useGPT4V}
                        gpt4vInput={gpt4vInput}
                        vectorFields={vectorFields}
                        showSemanticRankerOption={showSemanticRankerOption}
                        showQueryRewritingOption={showQueryRewritingOption}
                        showReasoningEffortOption={showReasoningEffortOption}
                        showGPT4VOptions={showGPT4VOptions}
                        showVectorOption={showVectorOption}
                        useOidSecurityFilter={useOidSecurityFilter}
                        useGroupsSecurityFilter={useGroupsSecurityFilter}
                        useLogin={!!useLogin}
                        loggedIn={loggedIn}
                        requireAccessControl={requireAccessControl}
                        onChange={handleSettingsChange}
                        showAgenticRetrievalOption={showAgenticRetrievalOption}
                        useAgenticRetrieval={useAgenticRetrieval}
                    />
                </Panel>

                <div style={{ marginBottom: 16 }}>
                    <label style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }}>
                        <Icon iconName="Upload" style={{ marginRight: 6 }} />
                        <span style={{ fontWeight: 500 }}>Upload Filter Config:</span>
                        <input
                            type="file"
                            accept=".json,.docx"
                            onChange={handleConfigUpload}
                            style={{ marginLeft: 8 }}
                        />
                    </label>
                    {/* Removed 'Reset to Default' button as requested */}
                    {dynamicFilterConfig && (
                        <div style={{ color: "#0078d4", marginTop: 8 }}>
                            Custom filter config loaded.
                        </div>
                    )}
                </div>
            </div>
        </>
    );
}