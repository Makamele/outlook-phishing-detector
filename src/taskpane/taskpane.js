Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("analyze-btn").onclick = analyzeEmail;
    }
});

async function analyzeEmail() {
    const button = document.getElementById("analyze-btn");
    const resultsDiv = document.getElementById("analysis-results");
    
    // Show loading state
    button.disabled = true;
    button.textContent = "Analyzing...";
    resultsDiv.innerHTML = "<p>🔍 Analyzing email with AI...</p>";
    
    try {
        // Get email content and subject
        const emailData = await getEmailContent();
        
        // Send to your AI backend for analysis
        const analysis = await analyzeWithGeminiAI(emailData);
        
        displayResults(analysis);
    } catch (error) {
        console.error("Analysis error:", error);
        resultsDiv.innerHTML = `
            <div class="error">
                <h3>❌ Analysis Failed</h3>
                <p>Error: ${error.message}</p>
            </div>
        `;
    } finally {
        button.disabled = false;
        button.textContent = "Analyze Email";
    }
}

function getEmailContent() {
    return new Promise((resolve, reject) => {
        const item = Office.context.mailbox.item;
        
        // Get both body and subject
        item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
            if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
                const emailData = {
                    subject: item.subject,
                    body: bodyResult.value,
                    sender: item.from ? item.from.displayName + " <" + item.from.emailAddress + ">" : "Unknown",
                    timestamp: item.dateTimeCreated ? item.dateTimeCreated.toISOString() : new Date().toISOString()
                };
                resolve(emailData);
            } else {
                reject(new Error("Failed to get email content"));
            }
        });
    });
}

async function analyzeWithGeminiAI(emailData) {
    // Your AI backend endpoint (we'll create this)
    const API_ENDPOINT = "https://your-domain.netlify.app/.netlify/functions/analyze-email";
    
    const response = await fetch(API_ENDPOINT, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            email_text: `Subject: ${emailData.subject}\nFrom: ${emailData.sender}\nBody: ${emailData.body}`
        })
    });
    
    if (!response.ok) {
        throw new Error(`API request failed: ${response.status}`);
    }
    
    const result = await response.json();
    return parseGeminiResponse(result.analysis);
}

function parseGeminiResponse(geminiText) {
    // Parse the Gemini AI response to extract structured data
    const lines = geminiText.split('\n');
    let classification = 'Unknown';
    let explanation = '';
    let confidence = 0;
    
    for (const line of lines) {
        const lowerLine = line.toLowerCase();
        
        if (lowerLine.includes('classification:') || lowerLine.includes('result:')) {
            if (lowerLine.includes('phishing') || lowerLine.includes('suspicious')) {
                classification = 'Phishing';
            } else if (lowerLine.includes('legitimate') || lowerLine.includes('safe')) {
                classification = 'Legitimate';
            }
        }
        
        if (lowerLine.includes('confidence:') || lowerLine.includes('score:')) {
            const match = line.match(/(\d+)/);
            if (match) {
                confidence = parseInt(match[1]);
            }
        }
        
        if (lowerLine.includes('explanation:') || lowerLine.includes('reason:')) {
            explanation = line.split(':')[1]?.trim() || '';
        }
    }
    
    // If explanation is still empty, use the full response
    if (!explanation) {
        explanation = geminiText;
    }
    
    return {
        classification,
        explanation,
        confidence,
        risk: classification === 'Phishing' ? 'HIGH' : 'LOW',
        fullResponse: geminiText
    };
}

function displayResults(analysis) {
    const resultsDiv = document.getElementById("analysis-results");
    const isPhishing = analysis.classification === 'Phishing';
    
    resultsDiv.innerHTML = `
        <div class="analysis-result ${isPhishing ? 'phishing' : 'legitimate'}">
            <h3>${isPhishing ? '🚨' : '✅'} ${analysis.classification}</h3>
            
            <div class="confidence-bar">
                <label>Confidence: ${analysis.confidence}%</label>
                <div class="progress-bar">
                    <div class="progress-fill" style="width: ${analysis.confidence}%"></div>
                </div>
            </div>
            
            <div class="explanation">
                <h4>Analysis:</h4>
                <p>${analysis.explanation}</p>
            </div>
            
            <details>
                <summary>Full AI Response</summary>
                <pre>${analysis.fullResponse}</pre>
            </details>
        </div>
    `;
}