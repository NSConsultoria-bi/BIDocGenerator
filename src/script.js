
"use strict";

document.addEventListener("DOMContentLoaded", () => {
    // =========================================
    // CONFIGURAÇÕES
    // =========================================

    const API_BASE_URL = "http://127.0.0.1:8000";
    const MAX_FILE_SIZE = 200 * 1024 * 1024;
    

    let selectedFile = null;
    let generatedDocx = null;
    let backendConnected = false;
    let isUploading = false;

    // =========================================
    // ELEMENTOS DO HTML
    // =========================================

    const fileInput = document.getElementById("project-file");
    const selectedFileContainer = document.getElementById("selected-file");
    const selectedFileName = document.getElementById("selected-file-name");
    const selectedFileSize = document.getElementById("selected-file-size");
    const removeFileButton = document.getElementById("remove-file");
    const generateButton = document.getElementById("generate-button");
    const processingNote = document.getElementById("processing-note");

    const downloadOptions = document.getElementById("download-options");
    const downloadFreeButton = document.getElementById("download-free-button"); 

    // =========================================
    // VALIDAR ELEMENTOS
    // =========================================

    if (
        !fileInput ||
        !selectedFileContainer ||
        !selectedFileName ||
        !selectedFileSize ||
        !removeFileButton ||
        !generateButton ||
        !processingNote ||
        !downloadOptions ||
        !downloadFreeButton 
        
    ) {
        console.error(
            "E-BI DOC Generator: elementos obrigatórios não encontrados."
        );
        return;
    }

    // =========================================
    // FUNÇÕES AUXILIARES
    // =========================================

    function formatFileSize(bytes) {
        if (bytes < 1024) {
            return `${bytes} bytes`;
        }

        if (bytes < 1024 * 1024) {
            return `${(bytes / 1024).toFixed(1)} KB`;
        }

        return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
    }

    function setStatus(message) {
        processingNote.textContent = message;
    }

    function updateGenerateButton() {
        generateButton.disabled =
            !selectedFile || !backendConnected || isUploading;
    }

    function resetDownload() {
        generatedDocx = null;
        downloadOptions.hidden = true;
    }

    function resetSelection(
        message = "Selecione um arquivo ZIP para começar."
    ) {
        selectedFile = null;
        fileInput.value = "";

        selectedFileContainer.hidden = true;
        selectedFileName.textContent = "";
        selectedFileSize.textContent = "";

        resetDownload();
        updateGenerateButton();
        setStatus(message);
    }

    function getSafeProjectName(filename) {
        return filename
            .replace(/\.zip$/i, "")
            .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
            .trim() || "power_bi";
    }

    // =========================================
    // CONEXÃO COM O BACKEND
    // =========================================

    async function checkBackend() {
        backendConnected = false;
        updateGenerateButton();

        try {
            const response = await fetch(`${API_BASE_URL}/health`);

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}`);
            }

            const data = await response.json();

            if (data.status !== "ok") {
                throw new Error("Resposta inesperada do backend.");
            }

            backendConnected = true;
            updateGenerateButton();

            setStatus(
                selectedFile
                    ? "Backend conectado. Projeto pronto para processamento."
                    : "Sistema conectado. Selecione um arquivo ZIP do projeto Power BI."
            );

            console.info("Backend conectado:", data);
        } catch (error) {
            backendConnected = false;
            updateGenerateButton();

            console.error("Falha na conexão com o backend:", error);

            setStatus(
                "Backend indisponível. Verifique se o servidor Python está em execução."
            );
        }
    }

    // =========================================
    // SELEÇÃO DO PROJETO
    // =========================================

    function handleFileSelection() {
        if (isUploading) {
            return;
        }

        const file = fileInput.files?.[0];

        if (!file) {
            resetSelection();
            return;
        }

        if (!file.name.toLowerCase().endsWith(".zip")) {
            resetSelection(
                "Formato inválido. Selecione um arquivo ZIP."
            );
            return;
        }

        if (file.size === 0) {
            resetSelection("O arquivo selecionado está vazio.");
            return;
        }

        if (file.size > MAX_FILE_SIZE) {
            resetSelection(
                "O arquivo excede o limite de 200 MB."
            );
            return;
        }

        selectedFile = file;
        resetDownload();

        selectedFileName.textContent = file.name;
        selectedFileSize.textContent = formatFileSize(file.size);
        selectedFileContainer.hidden = false;

        updateGenerateButton();

        setStatus(
            backendConnected
                ? "Arquivo selecionado. Clique em Analisar projeto."
                : "Arquivo selecionado. Aguardando conexão com o backend."
        );
    }

    fileInput.addEventListener("change", handleFileSelection);

    // =========================================
    // REMOVER PROJETO
    // =========================================

    removeFileButton.addEventListener("click", () => {
        if (isUploading) {
            return;
        }

        resetSelection(
            backendConnected
                ? "Arquivo removido. Selecione outro projeto."
                : "Arquivo removido. O backend está indisponível."
        );
    });

    // =========================================
    // ENVIAR PROJETO E RECEBER DOCX
    // =========================================

    async function uploadProject(file) {
        const formData = new FormData();
        formData.append("file", file, file.name);

        const response = await fetch(
            `${API_BASE_URL}/upload-project`,
            {
                method: "POST",
                body: formData
            }
        );

        if (!response.ok) {
            let message = `Erro HTTP ${response.status}`;

            try {
                const data = await response.json();

                if (typeof data.detail === "string") {
                    message = data.detail;
                } else if (data.detail?.message) {
                    message = data.detail.message;

                    if (Array.isArray(data.detail.errors)) {
                        message += ` ${data.detail.errors.join("; ")}`;
                    }
                }
            } catch {
                // Mantém a mensagem HTTP original.
            }

            throw new Error(message);
        }

        const contentType = response.headers.get("content-type") || "";

        if (
            !contentType.includes(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        ) {
            throw new Error(
                "O backend não retornou um documento DOCX válido."
            );
        }

        const blob = await response.blob();

        if (blob.size === 0) {
            throw new Error("O documento gerado está vazio.");
        }

        return blob;
    }

    // =========================================
    // DOWNLOAD DO DOCX
    // =========================================

    function downloadDocx() {
        if (!generatedDocx) {
            setStatus("Nenhum documento disponível para download.");
            return;
        }

        const filename =
            `documentacao_${getSafeProjectName(selectedFile.name)}.docx`;

        const objectUrl = URL.createObjectURL(generatedDocx);
        const link = document.createElement("a");

        link.href = objectUrl;
        link.download = filename;
        link.style.display = "none";

        document.body.appendChild(link);
        link.click();
        link.remove();

        window.setTimeout(() => {
            URL.revokeObjectURL(objectUrl);
        }, 1000);
    }

    // =========================================
    // OPÇÃO GRATUITA
    // =========================================

    
downloadFreeButton.addEventListener("click", () => {
    if (!generatedDocx || !selectedFile) {
        setStatus("Gere o documento antes de iniciar o download.");
        return;
    }

    // Abrir a página de contribuição em uma nova aba.
    const paymentWindow = window.open(
        "https://link.mercadopago.com.br/vertecnologias",
        "_blank",
        "noopener,noreferrer"
    );

    // Iniciar o download independentemente da abertura da nova aba.
    downloadDocx();

    if (paymentWindow === null) {
        setStatus(
            "Download iniciado. A nova aba pode ter sido bloqueada pelo navegador."
        );
    } else {
        setStatus(
            "Download iniciado. A contribuição pelo Mercado Pago é opcional."
        );
    }
});

    
        

    // =========================================
    // GERAR DOCUMENTAÇÃO
    // =========================================

    generateButton.addEventListener("click", async () => {
        if (isUploading) {
            return;
        }

        if (!selectedFile) {
            setStatus("Selecione um arquivo ZIP válido.");
            return;
        }

        if (!backendConnected) {
            setStatus("Backend indisponível. Verificando conexão...");
            await checkBackend();
            return;
        }

        const fileToUpload = selectedFile;

        isUploading = true;
        resetDownload();
        updateGenerateButton();

        generateButton.textContent = "Analisando projeto...";
        setStatus(
            "Enviando o projeto. A análise e a geração do DOCX podem levar alguns minutos."
        );

        try {
            generatedDocx = await uploadProject(fileToUpload);

            downloadOptions.hidden = false;

            setStatus(
                `Documento gerado com sucesso (${formatFileSize(generatedDocx.size)}). O download está disponível gratuitamente.`
            );

            downloadOptions.scrollIntoView({
                behavior: "smooth",
                block: "center"
            });

            console.info("DOCX gerado:", {
                size: generatedDocx.size,
                type: generatedDocx.type
            });
        } catch (error) {
            generatedDocx = null;
            downloadOptions.hidden = true;

            console.error("Erro no processamento:", error);

            setStatus(`Falha no processamento: ${error.message}`);
        } finally {
            isUploading = false;

            generateButton.innerHTML =
                'Analisar projeto <span aria-hidden="true">→</span>';

            updateGenerateButton();
        }
    });

    // =========================================
    // INICIALIZAÇÃO
    // =========================================

    resetSelection();
    checkBackend();
});
