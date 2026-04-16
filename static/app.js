document.addEventListener('DOMContentLoaded', () => {
    const submitBtn = document.getElementById('submitBtn');
    const queryInput = document.getElementById('queryInput');
    const resultsPanel = document.getElementById('resultsPanel');
    const finalResponse = document.getElementById('finalResponse');
    const scoreText = document.getElementById('scoreText');
    const scoreCircle = document.querySelector('.score-circle');
    const hallucinationAlert = document.getElementById('hallucinationAlert');
    const checkerLogsSection = document.getElementById('checkerLogsSection');
    const checkerLogsList = document.getElementById('checkerLogsList');

    const nodes = [
        document.getElementById('node-retrieve'),
        document.getElementById('node-solver'),
        document.getElementById('node-proposer'),
        document.getElementById('node-checker'),
        document.getElementById('node-synth')
    ];

    const sampleQuestions = document.getElementById('sampleQuestions');
    if (sampleQuestions) {
        sampleQuestions.addEventListener('change', (e) => {
            if (e.target.value) {
                queryInput.value = e.target.value;
            }
        });
    }

    function resetUI() {
        resultsPanel.style.display = 'none';
        checkerLogsSection.style.display = 'none';
        checkerLogsList.innerHTML = '';
        nodes.forEach(n => n.classList.remove('active'));
    }

    async function simulateNodeProgress() {
        // Just for visual effect to show the agents working
        for (let i = 0; i < nodes.length; i++) {
            nodes[i].classList.add('active');
            await new Promise(r => setTimeout(r, 600)); // wait 600ms per node
        }
    }

    submitBtn.addEventListener('click', async () => {
        const query = queryInput.value.trim();
        if (!query) return;

        resetUI();
        submitBtn.disabled = true;
        submitBtn.innerText = 'Processing Workflow...';

        // Start visual simulation
        simulateNodeProgress();

        try {
            const response = await fetch('/ask', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ query })
            });

            if (!response.ok) {
                throw new Error('API Request Failed');
            }

            const data = await response.json();
            
            // Populate Results
            resultsPanel.style.display = 'block';
            finalResponse.innerText = data.final_response;

            const score = data.confidence_score * 100;
            scoreText.innerText = `${score}%`;

            if (data.hallucinations_found) {
                scoreCircle.classList.add('danger');
                hallucinationAlert.className = 'alert danger';
                hallucinationAlert.innerText = '⚠️ Hallucinations Detected & Scrubbed!';
                
                checkerLogsSection.style.display = 'block';
                data.checker_results.forEach(res => {
                    if (res.is_hallucination) {
                        const li = document.createElement('li');
                        li.innerHTML = `
                            <p><strong>Claim:</strong> "${res.proposer_answer}"</p>
                            <p><strong>Database Verdict:</strong> ${res.evidence}</p>
                        `;
                        checkerLogsList.appendChild(li);
                    }
                });
            } else {
                scoreCircle.classList.remove('danger');
                hallucinationAlert.className = 'alert success';
                hallucinationAlert.innerText = '✅ 100% Factually Grounded';
            }

            // Ensure all nodes light up upon completion
            nodes.forEach(n => n.classList.add('active'));

        } catch (error) {
            alert('Error processing request. Ensure Neo4j and Ollama are running.');
            console.error(error);
        } finally {
            submitBtn.disabled = false;
            submitBtn.innerText = 'Execute Workflow';
        }
    });
});
