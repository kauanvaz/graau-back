# API de Geração de Relatórios - Documentação Atualizada

Esta API permite a geração assíncrona de relatórios DOCX a partir de dados do SharePoint, com gerenciamento de armazenamento de arquivos, processamento assíncrono e limpeza automática.

## Endpoints

### 1. Upload de imagem de capa

**Endpoint:** `POST /api/upload-cover-image`

**Descrição:** Faz upload de uma imagem para ser usada como capa no relatório.

**Corpo da Requisição:** Multipart form data com campo `file` contendo a imagem.

**Limitações:**
- Formatos permitidos: png, jpg, jpeg
- Tamanho máximo: 5MB

**Resposta (201 Created):**
```json
{
  "success": true,
  "message": "Imagem de capa enviada com sucesso",
  "image_id": "f7e9d2c1-b3a5-4e8f-9c6d-0b2a1e3f4d5c",
  "filename": "cover_f7e9d2c1-b3a5-4e8f-9c6d-0b2a1e3f4d5c.jpg"
}
```

**Resposta (400 Bad Request):**
```json
{
  "error": "Nenhum arquivo enviado"
}
```
ou
```json
{
  "error": "Formato de arquivo não permitido. Use: png, jpg, jpeg"
}
```
ou
```json
{
  "error": "Arquivo muito grande. Tamanho máximo: 5MB"
}
```

### 2. Consultar dados do SharePoint

**Endpoint:** `GET /api/sharepoint_data/<sharepoint_id>`

**Descrição:** Recupera dados do SharePoint com base no ID fornecido.

**Parâmetros de URL:**
- `sharepoint_id`: ID do item no SharePoint 

**Resposta (202 Accepted):**
```json
{
  "divisao_origem_ajustada": "DFPESSOAL/DFPESSOAL1",
  "divisao_origem_ajustada_diretoria": "Diretoria de Fiscalização de Pessoal e Previdência",
  "divisao_origem_ajustada_divisao": "1ª Divisão Técnica de Fiscalização",
  "equipe_fiscalizacao": ["Fulaxo X", "Fulano Y"],
  "exercicios": ["2024", "2025"],
  "n_processo_eTCE": "TC/00000/2025",
  "processo_tipo": "...-...",
  "procurador": "Fulano",
  "relator": "Sicrano",
  "subclasse": "DENÚNCIA",
  "unidades_fiscalizadas": ["P. M. DE X", "P. M. DE Y"],
  "VRF": "R$ 00,00",
}
```

### 3. Gerar Relatório (Assíncrono)

**Endpoint:** `POST /api/generate-report`

**Descrição:** Inicia a geração assíncrona de um novo relatório baseado nos parâmetros fornecidos. `report_params` deve conter tanto os valores recuperados do Sharepoint como os demais provenientes exclusivamente do frontend (passados pelo usuário). Nomes a combinar...

**Body (JSON):**
```json
{
  "report_params": {
    "divisao_origem_ajustada": "DFPESSOAL/DFPESSOAL1",
    "divisao_origem_ajustada_diretoria": "Diretoria de Fiscalização de Pessoal e Previdência",
    "divisao_origem_ajustada_divisao": "1ª Divisão Técnica de Fiscalização",
    "equipe_fiscalizacao": ["Fulaxo X", "Fulano Y"],
    "exercicios": ["2024", "2025"],
    "n_processo_eTCE": "TC/00000/2025",
    "processo_tipo": "...-...",
    "procurador": "Fulano",
    "relator": "Sicrano",
    "subclasse": "DENÚNCIA",
    "unidades_fiscalizadas": ["P. M. DE X", "P. M. DE Y"],
    "VRF": "R$ 00,00",
    "front1": "...",
    "front2": "...",
    "front3": "..."
  },
  "cover_image_id": "f7e9d2c1-b3a5-4e8f-9c6d-0b2a1e3f4d5c"
}
```

**Campos Obrigatórios:**
- `report_params`: Parâmetros para personalização do relatório
- `cover_image_id`: ID da imagem de capa previamente enviada

**Resposta (202 Accepted):**
```json
{
  "success": true,
  "message": "Geração de relatório iniciada",
  "task_id": "f7e9d2c1-b3a5-4e8f-9c6d-0b2a1e3f4d5c",
  "status": "processing"
}
```

**Resposta (400 Bad Request):**
```json
{
  "error": "Campo(s) obrigatório(s) ausente(s): report_params, cover_image_id"
}
```

**Resposta (404 Not Found):**
```json
{
  "error": "Imagem de capa não encontrada. Faça o upload novamente."
}
```

### 4. Verificar Status do Relatório

**Endpoint:** `GET /api/report-status/<task_id>`

**Descrição:** Verifica o status atual de uma tarefa de geração de relatório.

**Parâmetros de URL:**
- `task_id`: ID único da tarefa de geração

**Resposta (200 OK) - Em Processamento:**
```json
{
  "status": "processing",
  "message": "Gerando relatório com os dados obtidos",
  "progress": 50,
  "created_at": "2023-02-15T12:30:45"
}
```

**Resposta (200 OK) - Concluído:**
```json
{
  "status": "completed",
  "message": "Relatório gerado com sucesso",
  "progress": 100,
  "created_at": "2023-02-15T12:31:15",
  "download_url": "/api/reports/f7e9d2c1-b3a5-4e8f-9c6d-0b2a1e3f4d5c",
  "filename": "relatorio_20230215_123045_f7e9d2c1.docx"
}
```

**Resposta (200 OK) - Erro:**
```json
{
  "status": "error",
  "message": "Erro ao gerar relatório: Falha ao conectar ao SharePoint",
  "progress": 0,
  "created_at": "2023-02-15T12:30:45"
}
```

**Resposta (404 Not Found):**
```json
{
  "error": "Tarefa não encontrada"
}
```

### 5. Download de Relatório

**Endpoint:** `GET /api/reports/<report_id>`

**Descrição:** Faz o download de um relatório específico.

**Parâmetros de URL:**
- `report_id`: ID da tarefa (task_id) ou ID que aparece no nome do arquivo do relatório

**Resposta (200 OK):** Arquivo DOCX para download

**Resposta (404 Not Found):**
```json
{
  "error": "Relatório não encontrado"
}
```
ou
```json
{
  "error": "Arquivo de relatório não encontrado no servidor"
}
```

## Configurações do Sistema

A API possui as seguintes configurações:

1. **REPORT_EXPIRATION_MINUTES**: Tempo (em minutos) que os relatórios permanecerão no sistema antes de serem excluídos automaticamente. Valor atual: 15 minutos.

2. **CLEANUP_INTERVAL_SECONDS**: Intervalo (em segundos) entre execuções da rotina de limpeza de relatórios antigos. Valor atual: 300 segundos (5 minutos).

3. **ALLOWED_IMAGE_EXTENSIONS**: Extensões de arquivo permitidas para imagens de capa: png, jpg, jpeg.

4. **MAX_IMAGE_SIZE**: Tamanho máximo permitido para arquivos de imagem de capa: 5MB.

## Fluxo de Geração Assíncrona

1. Cliente faz upload da imagem de capa
2. Cliente consulta dados do SharePoint.
3. Cliente envia solicitação para gerar relatório incluindo `report_params` e `cover_image_id`
4. API responde imediatamente com um `task_id`
5. Cliente verifica o status da tarefa periodicamente
6. Quando o relatório estiver pronto, o cliente recebe um link para download
7. O relatório permanece disponível pelo período definido em REPORT_EXPIRATION_MINUTES

## Exemplos de Uso

### Upload de imagem de capa e geração de relatório

```javascript
// Exemplo de upload de imagem de capa
async function uploadCoverImage(fileInput) {
  const formData = new FormData();
  formData.append('file', fileInput.files[0]);
  
  const response = await fetch('/api/upload-cover-image', {
    method: 'POST',
    body: formData
  });
  
  const data = await response.json();
  
  if (data.success) {
    // Armazenar o ID da imagem para uso posterior
    const imageId = data.image_id;
    return imageId;
  } else {
    console.error("Erro ao fazer upload da imagem:", data.error);
    return null;
  }
}

// Consultar dados do SharePoint
async function getSharePointData(sharePointId) {
  const response = await fetch(`/api/sharepoint_data/${sharePointId}`);
  return await response.json();
}

// Exemplo de geração de relatório
async function generateReport(imageId) {
  // Iniciar a geração do relatório
  const response = await fetch('/api/generate-report', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      report_params: {
        divisao_origem_ajustada: "DFPESSOAL/DFPESSOAL1",
        divisao_origem_ajustada_diretoria: "Diretoria de Fiscalização de Pessoal e Previdência",
        divisao_origem_ajustada_divisao: "1ª Divisão Técnica de Fiscalização",
        equipe_fiscalizacao: ["Fulaxo X", "Fulano Y"],
        exercicios: ["2024", "2025"],
        n_processo_eTCE: "TC/00000/2025",
        processo_tipo: "...-...",
        procurador: "Fulano",
        relator: "Sicrano",
        subclasse: "DENÚNCIA",
        unidades_fiscalizadas: ["P. M. DE X", "P. M. DE Y"],
        VRF: "R$ 00,00",
        front1: "...",
        front2: "...",
        front3: "..."
      },
      cover_image_id: imageId
    })
  });

  const data = await response.json();
  
  if (data.success) {
    const taskId = data.task_id;
    checkReportStatus(taskId);
  } else {
    console.error("Erro ao iniciar geração do relatório:", data.error);
  }
}

// Verificar status do relatório
async function checkReportStatus(taskId) {
  // Função para verificar o status
  const checkStatus = async () => {
    const response = await fetch(`/api/report-status/${taskId}`);
    const data = await response.json();
    
    if (data.status === "completed") {
      console.log("Relatório concluído! URL:", data.download_url);
      // Redirecionar para download ou mostrar link
      window.location.href = data.download_url;
      return;
    }
    
    if (data.status === "error") {
      console.error("Erro na geração:", data.message);
      return;
    }
    
    // Atualizar progresso na interface
    updateProgressBar(data.progress);
    
    // Verificar novamente em 2 segundos
    setTimeout(checkStatus, 2000);
  };
  
  // Iniciar verificação
  checkStatus();
}

// Fluxo completo
async function processReport() {
  // 1. Upload da imagem
  const fileInput = document.getElementById('coverImageInput');
  const imageId = await uploadCoverImage(fileInput);
  
  if (imageId) {
    // 2. Gerar relatório
    await generateReport(imageId);
  }
}
```