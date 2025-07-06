"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import { VisualFormattingSettingsModel } from "./settings";

const webkitSpeechRecognition = (window as any).webkitSpeechRecognition;

export class Visual implements IVisual {
    private target: HTMLElement;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    private readonly API_URL = "http://127.0.0.1:8044/query";
    private readonly FEEDBACK_URL = "http://127.0.0.1:8044/feedback";
    private isLoading: boolean = false;
    private inputBox: HTMLInputElement;
    private sendButton: HTMLButtonElement;

    private SAMPLE_QUESTIONS = [
        "Who are the top 3 reps with highest sales?",
        "Which Sales team had the highest profit in Q2 of 2022?",
        "What are the 3 cities suffering highest losses of total drug sales?"

    ];

    constructor(options: VisualConstructorOptions) {
        this.formattingSettingsService = new FormattingSettingsService();
        this.target = options.element;

        const chatbotHeading = document.createElement("div");
        chatbotHeading.className = "chatbot-heading";

        const chatbotIcon = document.createElement("img");
        chatbotIcon.src = require("../assets/image.png");
        chatbotIcon.alt = "Chatbot Icon";
        chatbotIcon.className = "chatbot-icon";

        const headingText = document.createElement("span");
        headingText.innerText = "Chatbot Assistance";

        chatbotHeading.appendChild(chatbotIcon);
        chatbotHeading.appendChild(headingText);
        this.target.insertBefore(chatbotHeading, this.target.firstChild);

        const wrapper = document.createElement("div");
        wrapper.className = "chat-wrapper";

        const container = document.createElement("div");
        container.className = "chat-container";

        const chatArea = document.createElement("div");
        chatArea.className = "chat-output";

        const preloadedHeading = document.createElement("div");
        preloadedHeading.className = "sample-heading";
        preloadedHeading.innerText = "Here are some of the things you can try";

        chatArea.appendChild(preloadedHeading);

        // Display each preloaded question like a clickable chat bubble
        this.SAMPLE_QUESTIONS.forEach((question) => {
            const questionMsg = document.createElement("div");
            questionMsg.className = "chat-message preloaded";
            questionMsg.innerText = question;

            questionMsg.onclick = () => {
                this.displayUserMessage(chatArea, question);
                this.fetchDataFromAPI(chatArea, question);
            };

            chatArea.appendChild(questionMsg);
        });


        const inputForm = document.createElement("div");
        inputForm.className = "chat-form";

        this.inputBox = document.createElement("input");
        this.inputBox.type = "text";
        this.inputBox.placeholder = "Send a message";
        this.inputBox.className = "chat-input";

        this.sendButton = document.createElement("button");
        this.sendButton.className = "chat-send-button";
        this.sendButton.innerHTML = `
            <svg class="send-icon" viewBox="0 0 24 24" fill="none">
                <path d="M2 21L23 12L2 3V10L17 12L2 14V21Z" fill="#fff"/>
            </svg>
        `;

        const micButton = document.createElement("button");
        micButton.className = "chat-mic-button";
        micButton.innerHTML = `
            <svg viewBox="0 0 24 24" width="20" height="20">
                <path fill="#fff" d="M12 14a3 3 0 003-3V5a3 3 0 00-6 0v6a3 3 0 003 3zM19 11a1 1 0 00-2 0 5 5 0 01-10 0 1 1 0 00-2 0 7 7 0 006 6.93V21h2v-3.07A7 7 0 0019 11z"/>
            </svg>
        `;
        micButton.onclick = () => this.startVoiceInput();

        this.sendButton.addEventListener("click", () => this.handleSubmit(chatArea, preloadedHeading));
        this.inputBox.addEventListener("keydown", (e) => {
            if (e.key === "Enter") this.handleSubmit(chatArea, preloadedHeading);
        });

        inputForm.appendChild(this.inputBox);
        inputForm.appendChild(this.sendButton);
        inputForm.appendChild(micButton);
        container.appendChild(chatArea);
        container.appendChild(inputForm);
        wrapper.appendChild(container);
        this.target.appendChild(wrapper);
    }

    private handleSubmit(chatArea: HTMLElement, sampleBox: HTMLElement) {
        const userMessage = this.inputBox.value.trim();
        if (!userMessage || this.isLoading) return;

        this.displayUserMessage(chatArea, userMessage);
        this.inputBox.value = "";
        sampleBox.style.display = "none";
        this.fetchDataFromAPI(chatArea, userMessage);
    }

    private displayUserMessage(chatArea: HTMLElement, message: string) {
        const userMsg = document.createElement("div");
        userMsg.className = "chat-message user";
        userMsg.innerText = message;
        chatArea.appendChild(userMsg);
        chatArea.scrollTop = chatArea.scrollHeight;
    }

    private displayBotLoader(chatArea: HTMLElement): HTMLElement {
        const loader = document.createElement("div");
        loader.className = "chat-message bot loading";
        loader.innerHTML = `<span class="spinner"></span> Chatbot is typing...`;
        chatArea.appendChild(loader);
        chatArea.scrollTop = chatArea.scrollHeight;
        return loader;
    }

    private displayBotResponse(chatArea: HTMLElement, loader: HTMLElement, data: any) {
        loader.remove();
        const botMsg = document.createElement("div");
        botMsg.className = "chat-message bot";

        if (Array.isArray(data) && data.length > 1) {
            botMsg.innerHTML = data[1];
        } else {
            botMsg.innerText = "Invalid response format.";
        }

        const readoutButton = document.createElement("button");
        readoutButton.className = "readout-button";
        readoutButton.title = "Read aloud";
        readoutButton.innerHTML = `
            <svg viewBox="0 0 24 24" width="18" height="18">
                <path fill="#000" d="M3 10v4h4l5 5V5l-5 5H3zm13.5 2a2.5 2.5 0 010-5v2a.5.5 0 000 1v2zm2.5 0a5 5 0 000-10v2a3 3 0 010 6v2zm2.5 0a7.5 7.5 0 000-15v2a5.5 5.5 0 010 11v2z"/>
            </svg>
        `;

        let selectedVoice: SpeechSynthesisVoice | null = null;

        function loadVoicesAndSpeak(text: string) {
            const voices = speechSynthesis.getVoices();
            console.log("Voices loaded:", voices);
            
            // Try to find a female voice
            const femaleVoice = voices.find(v =>
                v.name.toLowerCase().includes("female") ||
                v.name.toLowerCase().includes("zira") ||
                v.name.toLowerCase().includes("aria") ||
                v.name.toLowerCase().includes("jenny") ||
                v.name.toLowerCase().includes("elsa")
            );

            // Fallback to male voice
            const maleVoice = voices.find(v =>
                v.name.toLowerCase().includes("male") ||
                v.name.toLowerCase().includes("guy") ||
                v.name.toLowerCase().includes("david") ||
                v.name.toLowerCase().includes("mark")
            );
        
            const utterance = new SpeechSynthesisUtterance(text);
            if (femaleVoice) {
                utterance.voice = femaleVoice;
                speechSynthesis.speak(utterance);
                console.log("Using female voice:", femaleVoice.name);
            } else if (maleVoice) {
                utterance.voice = maleVoice;
                speechSynthesis.speak(utterance);
                console.log("Using male voice:", maleVoice.name);
            } else {
                console.warn("No preferred voice found. Using default.");
            }
        }
        readoutButton.onclick = () => {
            if (speechSynthesis.getVoices().length > 0) {
                loadVoicesAndSpeak(botMsg.innerText);
            } else {
                speechSynthesis.onvoiceschanged = () => {
                    loadVoicesAndSpeak(botMsg.innerText);
                };
            }
        };
        botMsg.appendChild(readoutButton);
        
        // === Feedback Section ===
        const feedbackContainer = document.createElement("div");
        feedbackContainer.className = "feedback-container";

        const thumbsUp = document.createElement("button");
        thumbsUp.className = "thumb-button thumbs-up";
        thumbsUp.innerHTML = "👍";

        const thumbsDown = document.createElement("button");
        thumbsDown.className = "thumb-button thumbs-down";
        thumbsDown.innerHTML = "👎";

        const feedbackBox = document.createElement("textarea");
        feedbackBox.className = "feedback-textbox";
        feedbackBox.placeholder = "Optional feedback...";
        feedbackBox.style.display = "none";

        const submitButton = document.createElement("button");
        submitButton.className = "submit-feedback-button";
        submitButton.innerText = "Submit Feedback";
        submitButton.style.display = "none";

        let feedbackType = "";

        // === Click Handlers ===
        thumbsUp.onclick = () => {
            feedbackType = "positive";
            feedbackBox.style.display = "block";
            feedbackBox.placeholder = "Any feedback you'd like to share? (Optional)";
            submitButton.style.display = "inline-block";
        };

        thumbsDown.onclick = () => {
            feedbackType = "negative";
            feedbackBox.style.display = "block";
            feedbackBox.placeholder = "Tell us what went wrong (Required)";
            submitButton.style.display = "inline-block";
        };

        submitButton.onclick = () => {
            const text = feedbackBox.value.trim();
            if (feedbackType === "negative" && text === "") {
                alert("Please provide feedback for thumbs down.");
                return;
            }

            // === Send Feedback to Python API ===
            fetch(this.FEEDBACK_URL, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    feedbackType: feedbackType,
                    feedbackText: text
                })
            }).then(res => {
                if (res.ok) {
                    feedbackContainer.innerHTML = "✅ Feedback submitted. Thanks!";
                } else {
                    feedbackContainer.innerHTML = "⚠️ Failed to submit feedback.";
                }
            }).catch(() => {
                feedbackContainer.innerHTML = "⚠️ Error sending feedback.";
            });
        };

        // === Append to DOM ===
        feedbackContainer.appendChild(thumbsUp);
        feedbackContainer.appendChild(thumbsDown);
        feedbackContainer.appendChild(feedbackBox);
        feedbackContainer.appendChild(submitButton);
        botMsg.appendChild(feedbackContainer);


        chatArea.appendChild(botMsg);
        chatArea.scrollTop = chatArea.scrollHeight;
    }

    private async fetchDataFromAPI(chatArea: HTMLElement, userMessage: string) {
        this.isLoading = true;
        this.inputBox.disabled = true;
        this.sendButton.disabled = true;

        const loader = this.displayBotLoader(chatArea);

        try {
            const response = await fetch(this.API_URL, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ userMessage })
            });

            const raw = await response.text();
            console.log("🚀 Raw response from API:", raw);

            let data: any;
            try {
                data = JSON.parse(raw);
            } catch (jsonErr) {
                console.error("❌ JSON parse error:", jsonErr);
                loader.innerText = "⚠️ Server returned invalid JSON.";
                return;
            }

            this.displayBotResponse(chatArea, loader, data);
        } catch (error) {
            console.error("❌ Fetch error:", error);
            loader.innerText = "⚠️ Network error. Check if the API is running.";
        } finally {
            this.isLoading = false;
            this.inputBox.disabled = false;
            this.sendButton.disabled = false;
        }
    }


    private startVoiceInput() {
        if (!webkitSpeechRecognition) {
            alert("Speech recognition is not supported in this browser.");
            return;
        }

        const recognition = new webkitSpeechRecognition();
        recognition.lang = "en-US";
        recognition.interimResults = false;
        recognition.maxAlternatives = 1;

        recognition.onresult = (event: any) => {
            const voiceText = event.results[0][0].transcript;
            this.inputBox.value = voiceText;
            this.sendButton.click();
        };

        recognition.onerror = (event: any) => {
            console.error("Speech recognition error:", event.error);
        };

        recognition.start();
    }

    public update(options: VisualUpdateOptions) {
        if (!options.dataViews || options.dataViews.length === 0) {
            return;
        }

        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
            VisualFormattingSettingsModel,
            options.dataViews[0]
        );
    }
}
