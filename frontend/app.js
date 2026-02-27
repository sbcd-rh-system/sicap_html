// ⚙️ URL do backend Python (Web Service no Render)
// Altere para a URL do seu Web Service após criar no Render
const API_BASE_URL = window.location.hostname === 'localhost'
    ? ''   // local: usa URL relativa (backend roda junto)
    : 'https://sicap-html-2.onrender.com'; // Web Service (backend Python)

document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('sicap-form');
    const fileInput = document.getElementById('file-upload');
    const uploadArea = document.getElementById('upload-area');
    const submitBtn = document.getElementById('submit-btn');
    const statusArea = document.getElementById('status-area');
    const fileNameDisplay = document.getElementById('file-name');
    const fileLabel = document.getElementById('file-label');
    const userInput = document.getElementById('sicap-user');
    const passInput = document.getElementById('sicap-pass');

    let isFileValid = false;

    // Drag & Drop
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        uploadArea.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        uploadArea.addEventListener(eventName, highlight, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        uploadArea.addEventListener(eventName, unhighlight, false);
    });

    function highlight() {
        uploadArea.classList.add('dragover');
    }

    function unhighlight() {
        uploadArea.classList.remove('dragover');
    }

    uploadArea.addEventListener('drop', handleDrop, false);

    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        handleFiles(files);
    }

    fileInput.addEventListener('change', function () {
        handleFiles(this.files);
    });

    // Validar form dinamicamente
    function checkFormValidity() {
        // Agora verificamos também se o form é válido nativamente (inputs required)
        // Mas o botão só habilita se arquivo estiver ok. 
        // HTML5 validação cuida do resto no submit.
        if (isFileValid) {
            submitBtn.disabled = false;
        } else {
            submitBtn.disabled = true;
        }
    }

    function handleFiles(files) {
        if (files.length > 0) {
            const file = files[0];
            if (validateFile(file)) {
                fileNameDisplay.textContent = file.name;
                fileLabel.textContent = "Arquivo selecionado:";
                document.querySelector('.file-info').style.display = 'block';
                isFileValid = true;
                checkFormValidity();
                submitBtn.classList.add('pulse');
            }
        }
    }

    function validateFile(file) {
        const validExtensions = ['.xlsx', '.xls'];
        const fileName = file.name.toLowerCase();
        const isValid = validExtensions.some(ext => fileName.endsWith(ext));

        if (!isValid) {
            showStatus('Apenas arquivos .xlsx ou .xls são permitidos.', 'error');
            fileInput.value = ''; // Clear input
            fileNameDisplay.textContent = '';
            isFileValid = false;
            checkFormValidity();
            return false;
        }
        hideStatus();
        return true;
    }

    function getSelectedRadio(name) {
        const radios = document.getElementsByName(name);
        for (let r of radios) {
            if (r.checked) return r.value;
        }
        return null;
    }

    // LISTENER NO FORM (SUBMIT)
    form.addEventListener('submit', async (e) => {
        e.preventDefault(); // Impede reload, mas navegador entende como submissão

        const file = fileInput.files[0];
        const user = userInput.value.trim();
        const pass = passInput.value.trim();
        const prestacaoIdManual = document.getElementById('prestacao-id').value.trim();

        if (!file) {
            showStatus('Selecione um arquivo Excel.', 'error');
            return;
        }
        // Validação HTML5 'required' já deve ter cuidado de user/pass/id, mas checamos por segurança
        if (!user || !pass || !prestacaoIdManual) {
            showStatus('Preencha todas as informações obrigatórias (Usuário, Senha e ID).', 'error');
            return;
        }

        // Reset UI
        submitBtn.disabled = true;
        submitBtn.classList.remove('pulse');
        showStatus('Autenticando e enviando... Isso pode levar alguns minutos.', 'loading');

        const formData = new FormData();
        formData.append('file', file);
        formData.append('usuario', user);
        formData.append('senha', pass);
        // mes e ano não são mais necessários para o envio manual
        formData.append('prestacao_id', prestacaoIdManual);

        try {
            const response = await fetch(`${API_BASE_URL}/api/processar`, {
                method: 'POST',
                body: formData
            });

            let result;
            const textResponse = await response.text();

            // Log de diagnóstico — ver no DevTools (F12 > Console)
            console.log('[SICAP] Status:', response.status);
            console.log('[SICAP] Resposta bruta:', textResponse);

            if (!textResponse || textResponse.trim() === '') {
                throw new Error(`Resposta vazia do servidor (Status ${response.status}). Verifique os logs do Render.`);
            }

            try {
                result = JSON.parse(textResponse);
            } catch (jsonErr) {
                throw new Error(`Resposta inválida (Status ${response.status}): ${textResponse.substring(0, 200)}`);
            }

            if (response.ok && result.status === 'sucesso') {
                showStatus(result.mensagem, 'success');
                console.log('Detalhes:', result.detalhes);

                const shouldSave = document.getElementById('save-credentials').checked;

                // Tenta salvar senha no navegador se a opção estiver marcada
                if (shouldSave && window.PasswordCredential) {
                    try {
                        const cred = new PasswordCredential({
                            id: user,
                            password: pass,
                            name: user,
                        });
                        await navigator.credentials.store(cred);
                    } catch (e) {
                        console.warn('Não foi possível solicitar o salvamento de senha via API:', e);
                    }
                }
            } else {
                const msg = result.mensagem || 'Erro desconhecido';
                const detalhes = result.detalhes ? JSON.stringify(result.detalhes, null, 2) : '';
                showStatus(`${msg} ${detalhes}`, 'error');
                console.error('Detalhes erro:', result);
            }

        } catch (error) {
            console.error('Erro:', error);
            showStatus(error.message || 'Falha ao conectar com o servidor.', 'error');
        } finally {
            submitBtn.disabled = false;
        }
    });

    function showStatus(msg, type) {
        statusArea.style.display = 'block';
        statusArea.className = '';
        statusArea.classList.add(`status-${type}`);

        if (msg.length > 300) {
            statusArea.style.overflow = "auto";
            statusArea.style.maxHeight = "200px";
        } else {
            statusArea.style.overflow = "visible";
            statusArea.style.maxHeight = "none";
        }

        if (type === 'loading') {
            statusArea.innerHTML = `<span class="spinner"></span> ${msg}`;
        } else {
            statusArea.innerText = msg;
        }
    }

    function hideStatus() {
        statusArea.style.display = 'none';
    }
});
