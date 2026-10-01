# 4U MetaViewer Pro 5.0 — Auditoria de Metadados, LGPD & Higienizador

> **4U MetaViewer Pro** é uma plataforma avançada de extração de metadados, análise forense digital, verificação de conformidade com LGPD/GDPR e **higienização de arquivos (Metadata Stripper/Sanitizer)**. Extraia dados profundos de imagens, documentos Office, planilhas, PDFs e mídias, visualize coordenadas GPS em um mapa interativo e limpe metadados antes de compartilhar arquivos confidenciais.

---

## 🌟 Principais Recursos

### 1. 🛡️ Higienizador de Metadados (Metadata Stripper / Sanitizer)
- **Remoção de 1 Clique**: Elimina dados EXIF, coordenadas de GPS, histórico de revisões, nomes de autores e marcas de câmeras.
- **Imagens (JPG, PNG, WebP)**: Re-renderiza e elimina todas as tags de metadados, comentários e miniaturas embutidas.
- **Documentos (DOCX, XLSX)**: Limpa os fluxos internos `core.xml`, `app.xml`, nomes de autores, empresas e comentários de revisão.
- **Download Seguro**: Gera um arquivo higienizado pronto para download imediato.

### 2. 🗺️ Geolocalização Real com Mapa Interativo (Leaflet / OpenStreetMap)
- Decodifica coordenadas GPS EXIF complexas (graus, minutos e segundos) para formato decimal preciso.
- Renderiza um mini-mapa interativo com marcador no local exato do registro fotográfico.
- Links diretos para **Google Maps**, **Waze** e **OpenStreetMap**.
- Alerta visual de segurança contra vazamento de localização física de residência ou escritório.

### 3. 🚦 Score de Risco de Privacidade & Diagnóstico LGPD
- **Medidor Percentual de Exposição (0 a 100%)**:
  - 🔴 **Alto Risco (Crítico)**: Exposição de GPS físico ou dados de identidade pessoal.
  - 🟡 **Atenção (Risco Médio)**: Organizações corporativas, softwares ou histórico de edições.
  - 🟢 **Baixo Risco (Seguro)**: Arquivo sem metadados sensíveis identificáveis.
- Checklist didático detalhando cada fator de risco detectado.

### 4. 🤖 Detecção de Criação por IA & Softwares de Edição
- Identificação de assinaturas e tags de ferramentas de inteligência artificial generativa:
  - *Midjourney, Stable Diffusion, DALL-E, Adobe Firefly, NovelAI, ComfyUI*.
- Identificação de softwares de edição:
  - *Adobe Photoshop, Lightroom, Canva, GIMP, CorelDRAW, Figma*.

### 5. 🔐 Assinatura Digital & Hashes Forenses
- Cálculo instantâneo de **MD5**, **SHA-1**, **SHA-256**, **SHA-512** e **CRC32**.
- Botões de cópia rápida para laudos periciais e auditorias de integridade.

### 6. 📑 Exportação de Laudos Periciais
- **Laudo Forense em PDF**: Formatação limpa pronta para impressão pericial (`window.print()`).
- **Exportação JSON**: Dados estruturados para integrações e scripts.
- **Laudo Markdown**: Cópia instantânea em formato markdown para relatórios técnicos.

### 7. ⚡ Demonstração em 1 Clique
- Permite testar o sistema sem a necessidade de upload imediato:
  - 📸 *Foto com GPS (iPhone 15 Pro na Av. Paulista / MASP)*
  - 📘 *Contrato Corporativo DOCX com histórico de edições*
  - 📊 *Planilha Financeira XLSX com fórmulas corporativas*

---

## 🔒 Privacidade & Segurança
Todos os arquivos enviados são analisados exclusivamente em memória e em diretórios temporários do servidor com descarte automático. Nenhum metadado é compartilhado com terceiros.

---

## 🛠️ Tecnologias
- **PHP 8.3 / 7.4+**
- **Smalot PDFParser**, **PhpOffice PhpWord**, **PhpOffice PhpSpreadsheet**, **getID3**
- **Leaflet.js & OpenStreetMap**
- **Tailwind CSS & 4U Glassmorphism Design System**

---

© 2026 [4U.IA.BR](https://4u.ia.br) — Todos os direitos reservados.
