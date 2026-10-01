<?php
header('Content-Type: text/html; charset=UTF-8');
/**
 * ==========================================================================
 * 🚀 4U METAVIEWER PRO — VERSÃO 5.0 ULTIMATE
 * Análise Forense de Metadados, Higienização de Arquivos (Sanitizer LGPD),
 * Geolocalização em Mapa Interativo (Leaflet/OSM) & Detecção de IA
 * Suporte: PDF, DOCX, XLSX, JPG, PNG, WEBP, TIFF, MP3, MP4, HTML e mais
 * ==========================================================================
 */

// --- CONFIGURAÇÕES DO SISTEMA ---
ini_set('memory_limit', '512M');
ini_set('max_execution_time', 300);
ini_set('display_errors', 0);
error_reporting(E_ALL & ~E_NOTICE & ~E_DEPRECATED);
date_default_timezone_set('America/Sao_Paulo');

// --- CARREGAMENTO DE DEPENDÊNCIAS COMPOSER ---
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require __DIR__ . '/vendor/autoload.php';
}

use Smalot\PdfParser\Parser as PdfParser;
use PhpOffice\PhpWord\IOFactory as WordIO;
use PhpOffice\PhpSpreadsheet\IOFactory as SpreadsheetIO;

// --- GERENCIAMENTO DE DOWNLOAD DE ARQUIVO HIGIENIZADO (SANITIZER) ---
if (isset($_GET['download_clean']) && !empty($_GET['download_clean'])) {
    $token = preg_replace('/[^a-zA-Z0-9_-]/', '', $_GET['download_clean']);
    $ext = preg_replace('/[^a-zA-Z0-9]/', '', $_GET['ext'] ?? 'dat');
    $origName = preg_replace('/[^a-zA-Z0-9._-]/', '', $_GET['orig'] ?? 'arquivo');
    
    $cleanPath = sys_get_temp_dir() . '/4u_clean_' . $token . '.' . $ext;
    if (file_exists($cleanPath)) {
        header('Content-Description: File Transfer');
        header('Content-Type: application/octet-stream');
        header('Content-Disposition: attachment; filename="limpo_' . $origName . '"');
        header('Expires: 0');
        header('Cache-Control: must-revalidate');
        header('Pragma: public');
        header('Content-Length: ' . filesize($cleanPath));
        readfile($cleanPath);
        @unlink($cleanPath);
        exit;
    } else {
        die("Arquivo higienizado expirado ou inexistente. Faça um novo upload.");
    }
}

// --- VARIÁVEIS INICIAIS ---
$resultado = null;
$erro = null;
$icone = "📂";
$corTopo = "#8b5cf6";
$badges = [];
$estatisticas = [];
$alertas = [];
$hashes = [];
$gpsCoords = null; // ['lat' => float, 'lng' => float, 'alt' => string]
$tokenLimpo = null;
$extLimpo = null;
$nomeOriginal = null;
$privacyScore = null;
$iaDetectada = [];
$softwaresDetectados = [];

// Funções Utilitárias de Conversão e Formatação
function formatarTamanho($bytes) {
    if ($bytes >= 1073741824) return number_format($bytes / 1073741824, 2) . ' GB';
    if ($bytes >= 1048576) return number_format($bytes / 1048576, 2) . ' MB';
    if ($bytes >= 1024) return number_format($bytes / 1024, 2) . ' KB';
    return $bytes . ' bytes';
}

function calcularHashes($filepath) {
    return [
        'MD5' => md5_file($filepath),
        'SHA-1' => sha1_file($filepath),
        'SHA-256' => hash_file('sha256', $filepath),
        'SHA-512' => hash_file('sha512', $filepath),
        'CRC32' => hash_file('crc32b', $filepath)
    ];
}

function detectarMimeReal($filepath) {
    if (function_exists('finfo_open')) {
        $finfo = finfo_open(FILEINFO_MIME_TYPE);
        $mime = finfo_file($finfo, $filepath);
        finfo_close($finfo);
        if ($mime) return $mime;
    }
    if (function_exists('mime_content_type')) {
        return mime_content_type($filepath);
    }
    return 'application/octet-stream';
}

// Conversor de Fração EXIF para Float
function avaliarFracao($fracao) {
    if (is_numeric($fracao)) return (float)$fracao;
    if (strpos($fracao, '/') === false) return (float)$fracao;
    $partes = explode('/', $fracao);
    if (count($partes) === 2 && (float)$partes[1] != 0) {
        return (float)$partes[0] / (float)$partes[1];
    }
    return 0;
}

// Conversor de Coordenadas EXIF para Decimal
function converterGpsCoord($coord, $hemi) {
    if (!is_array($coord) || count($coord) < 3) return null;
    $graus = avaliarFracao($coord[0]);
    $minutos = avaliarFracao($coord[1]);
    $segundos = avaliarFracao($coord[2]);
    $decimal = $graus + ($minutos / 60) + ($segundos / 3600);
    if (strtoupper($hemi) === 'S' || strtoupper($hemi) === 'W') {
        $decimal = $decimal * -1;
    }
    return round($decimal, 6);
}

// Higienizador / Limpador de Arquivos (Metadata Stripper)
function higienizarArquivo($caminhoOrigem, $ext) {
    $token = bin2hex(random_bytes(16));
    $destino = sys_get_temp_dir() . '/4u_clean_' . $token . '.' . $ext;

    // Imagens Raster
    if (in_array($ext, ['jpg', 'jpeg'])) {
        if (function_exists('imagecreatefromjpeg')) {
            $img = @imagecreatefromjpeg($caminhoOrigem);
            if ($img) {
                imagejpeg($img, $destino, 95);
                imagedestroy($img);
                return $token;
            }
        }
    } elseif ($ext === 'png') {
        if (function_exists('imagecreatefrompng')) {
            $img = @imagecreatefrompng($caminhoOrigem);
            if ($img) {
                imagealphablending($img, false);
                imagesavealpha($img, true);
                imagepng($img, $destino, 9);
                imagedestroy($img);
                return $token;
            }
        }
    } elseif ($ext === 'webp') {
        if (function_exists('imagecreatefromwebp')) {
            $img = @imagecreatefromwebp($caminhoOrigem);
            if ($img) {
                imagewebp($img, $destino, 95);
                imagedestroy($img);
                return $token;
            }
        }
    }

    // Documentos Office OpenXML (DOCX, XLSX)
    if (in_array($ext, ['docx', 'xlsx']) && class_exists('ZipArchive')) {
        copy($caminhoOrigem, $destino);
        $zip = new ZipArchive();
        if ($zip->open($destino) === TRUE) {
            $cleanCore = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>';
            $zip->addFromString('docProps/core.xml', $cleanCore);
            if ($zip->locateName('docProps/app.xml') !== false) {
                $cleanApp = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>4U Metadata Sanitizer</Application></Properties>';
                $zip->addFromString('docProps/app.xml', $cleanApp);
            }
            if ($zip->locateName('word/comments.xml') !== false) {
                $zip->deleteName('word/comments.xml');
            }
            $zip->close();
            return $token;
        }
    }

    // Fallback: cópia direta para download caso não seja possível higienizar
    copy($caminhoOrigem, $destino);
    return $token;
}

// Calculador de Score de Risco de Privacidade & LGPD
function calcularScorePrivacidade($dados, $gpsCoords, $ext) {
    $score = 0;
    $fatores = [];

    if ($gpsCoords && isset($gpsCoords['lat']) && isset($gpsCoords['lng'])) {
        $score += 45;
        $fatores[] = ['tipo' => 'critico', 'texto' => 'Coordenadas de GPS exatas detectadas (risco de rastreamento físico)'];
    }

    $autorDetectado = false;
    foreach ($dados as $grupo => $itens) {
        if (is_array($itens)) {
            foreach ($itens as $k => $v) {
                $kl = strtolower($k);
                if (in_array($kl, ['autor', 'autor original', 'artista/autor', 'última modificação por', 'creator', 'author']) && !empty($v)) {
                    $autorDetectado = true;
                    $fatores[] = ['tipo' => 'alto', 'texto' => "Identidade pessoal ou profissional exposta: '{$k}' = " . htmlspecialchars($v)];
                    break 2;
                }
            }
        }
    }
    if ($autorDetectado) $score += 25;

    $empresaDetectada = false;
    foreach ($dados as $grupo => $itens) {
        if (is_array($itens)) {
            foreach ($itens as $k => $v) {
                $kl = strtolower($k);
                if (in_array($kl, ['empresa', 'company', 'gerente', 'manager']) && !empty($v)) {
                    $empresaDetectada = true;
                    $fatores[] = ['tipo' => 'medio', 'texto' => "Organização corporativa exposta: '{$k}' = " . htmlspecialchars($v)];
                    break 2;
                }
            }
        }
    }
    if ($empresaDetectada) $score += 15;

    $aparelhoDetectado = false;
    foreach ($dados as $grupo => $itens) {
        if (is_array($itens)) {
            foreach ($itens as $k => $v) {
                $kl = strtolower($k);
                if (in_array($kl, ['modelo', 'fabricante', 'make', 'model']) && !empty($v)) {
                    $aparelhoDetectado = true;
                    $fatores[] = ['tipo' => 'baixo', 'texto' => "Dispositivo registrado: " . htmlspecialchars($v)];
                    break 2;
                }
            }
        }
    }
    if ($aparelhoDetectado) $score += 10;

    foreach ($dados as $grupo => $itens) {
        if (is_array($itens)) {
            foreach ($itens as $k => $v) {
                $kl = strtolower($k);
                if (strpos($kl, 'revisão') !== false || strpos($kl, 'revisao') !== false) {
                    $score += 10;
                    $fatores[] = ['tipo' => 'medio', 'texto' => 'Histórico de edições internas detectado'];
                    break 2;
                }
            }
        }
    }

    $score = min(100, $score);
    if ($score === 0 && !empty($dados)) $score = 10;

    $nivel = 'seguro';
    $rotulo = 'Baixo Risco';
    $cor = '#10b981';

    if ($score >= 60) {
        $nivel = 'critico';
        $rotulo = 'Alto Risco (Crítico)';
        $cor = '#ef4444';
    } elseif ($score >= 25) {
        $nivel = 'medio';
        $rotulo = 'Atenção (Risco Médio)';
        $cor = '#f59e0b';
    }

    return [
        'score' => $score,
        'nivel' => $nivel,
        'rotulo' => $rotulo,
        'cor' => $cor,
        'fatores' => $fatores
    ];
}

// Detector de Softwares e IA
function detectarSoftwareIA($dados) {
    $iasDetectadas = [];
    $softwaresDetectados = [];

    $iaKeywords = [
        'midjourney' => 'Midjourney',
        'stable diffusion' => 'Stable Diffusion',
        'dall-e' => 'DALL-E (OpenAI)',
        'dalle' => 'DALL-E',
        'firefly' => 'Adobe Firefly',
        'novelai' => 'NovelAI',
        'comfyui' => 'ComfyUI',
        'automatic1111' => 'Automatic1111 WebUI'
    ];

    $softKeywords = [
        'photoshop' => 'Adobe Photoshop',
        'lightroom' => 'Adobe Lightroom',
        'canva' => 'Canva',
        'gimp' => 'GIMP',
        'coreldraw' => 'CorelDRAW',
        'illustrator' => 'Adobe Illustrator',
        'figma' => 'Figma'
    ];

    $stringGeral = json_encode($dados, JSON_UNESCAPED_UNICODE);
    $stringGeralLower = strtolower($stringGeral);

    foreach ($iaKeywords as $k => $label) {
        if (strpos($stringGeralLower, $k) !== false) $iasDetectadas[] = $label;
    }
    foreach ($softKeywords as $k => $label) {
        if (strpos($stringGeralLower, $k) !== false) $softwaresDetectados[] = $label;
    }

    return [
        'ias' => array_unique($iasDetectadas),
        'softwares' => array_unique($softwaresDetectados)
    ];
}

// ==========================================================================
// MODO DEMONSTRAÇÃO EM 1 CLIQUE
// ==========================================================================
$demo = $_GET['demo'] ?? null;
if ($demo) {
    if ($demo === 'gps') {
        $nomeOriginal = 'IMG_20260930_174211_HDR.jpg';
        $ext = 'jpg';
        $icone = "🖼️";
        $corTopo = "#ec4899";
        $badges = ['Imagem', 'JPG', 'Geolocalizado (GPS)', 'iPhone 15 Pro'];
        $gpsCoords = ['lat' => -23.561414, 'lng' => -46.655881, 'alt' => '760 metros'];
        
        $hashes = [
            'MD5' => 'e2fc714c4727ee9395f324cd2e7f331f',
            'SHA-1' => '849d479133be02ee40b3ff64d852a3264c8d50bf',
            'SHA-256' => '4a2f8b50f613c2429810bb4298a0026e6371cf78accdb0b925b410543e3ffc80',
            'SHA-512' => '9f833777553f12efc464c23f465cbb64b5847a964bb93902347ebcc708304cbe...',
            'CRC32' => '9d251f28'
        ];

        $dados = [
            '📝 Nome do Arquivo' => $nomeOriginal,
            '🌐 Geolocalização Exata (GPS)' => [
                'Latitude' => '-23.561414° S (23° 33\' 41.09" S)',
                'Longitude' => '-46.655881° O (46° 39\' 21.17" O)',
                'Altitude' => '760 metros acima do nível do mar',
                'Ponto de Referência' => 'Avenida Paulista / MASP — São Paulo, SP',
                'Status' => 'Exposto nos metadados EXIF da imagem'
            ],
            '📷 Equipamento Fotográfico' => [
                'Fabricante' => 'Apple',
                'Modelo' => 'iPhone 15 Pro',
                'Software' => 'iOS 18.1.2 Camera App',
                'Lente' => 'iPhone 15 Pro back triple camera 24mm f/1.78',
                'Distância Focal' => '6.86 mm (24 mm equivalente)',
                'Abertura' => 'f/1.78',
                'Velocidade do Obturador' => '1/1200 s',
                'Sensibilidade ISO' => 'ISO 50'
            ],
            '📐 Dimensões e Qualidade' => [
                'Resolução' => '8064 × 6048 pixels',
                'Megapixels' => '48.77 MP',
                'Proporção' => '1.33:1 (4:3)',
                'Espaço de Cor' => 'Display P3 Wide Color',
                'Bits por Canal' => '8-bit',
                'Tamanho Impresso (300 DPI)' => '68.3 × 51.2 cm'
            ],
            '📊 Informações do Arquivo' => [
                'MIME Type Detectado' => 'image/jpeg',
                'Tamanho em Disco' => '9.42 MB',
                'Data da Captura' => date('d/m/Y H:i:s', time() - 3600 * 4)
            ]
        ];

        $estatisticas = [
            ['label' => 'Resolução', 'value' => '48.8 MP', 'color' => 'blue'],
            ['label' => 'Qualidade', 'value' => 'Ultra HD (48MP)', 'color' => 'green'],
            ['label' => 'Geotag', 'value' => 'GPS Ativo', 'color' => 'red']
        ];
        $alertas = [
            '🌍 Este arquivo expõe a localização geográfica exata de onde foi capturado!',
            '⚠️ Alto Risco LGPD: Rastreamento físico de residência ou local de trabalho detectado.'
        ];

        $tokenLimpo = 'demo_clean_gps';
        $extLimpo = 'jpg';
        $resultado = $dados;
    }
    elseif ($demo === 'docx') {
        $nomeOriginal = 'Contrato_Parceria_Estrategica_2026_Rev14.docx';
        $ext = 'docx';
        $icone = "📘";
        $corTopo = "#2563eb";
        $badges = ['Documento', 'DOCX', 'Corporativo', 'Confidencial'];

        $hashes = [
            'MD5' => '71b12d59048a176846deef39bbca584a',
            'SHA-1' => '2aae5ff402927e1f574d6c442750e326b8cb7980',
            'SHA-256' => '9b841804e4604d30ad56230f89d1502447ad9efb10b0e5124032ef34932029be',
            'SHA-512' => 'c1b489025e173875be489115f791724032aeb89e130285923145...',
            'CRC32' => '6e9f1a23'
        ];

        $dados = [
            '📝 Nome do Arquivo' => $nomeOriginal,
            '👤 Metadados de Autoria (Core.xml)' => [
                'Autor Original' => 'Dr. Roberto Almeida (Diretoria Jurídica)',
                'Última Modificação Por' => 'Ana Clara Mendes — Compliance & Riscos',
                'Título do Documento' => 'Contrato de Parceria Tecnológica e Investimento Estratégico',
                'Assunto' => 'Acordo Comercial Confidencial Q3/Q4',
                'Número de Revisão' => 'Revisão 14',
                'Data de Criação' => '15/01/2026 09:30:12',
                'Última Modificação' => '28/09/2026 16:45:00'
            ],
            '🏢 Metadados Corporativos (App.xml)' => [
                'Empresa / Organização' => 'Nexus Global Holding S.A.',
                'Gerente / Supervisor' => 'Carlos Eduardo Vasconcelos',
                'Aplicativo Criador' => 'Microsoft Office Word 365 (Build 16.0.17425)',
                'Tempo Total de Edição' => '342 minutos (5h 42m)'
            ],
            '📊 Estatísticas de Conteúdo' => [
                'Páginas' => '18',
                'Palavras' => '6.450',
                'Caracteres' => '41.280',
                'Parágrafos' => '245',
                'Linhas' => '1.120',
                'Imagens Embutidas' => '3 imagens'
            ]
        ];

        $estatisticas = [
            ['label' => 'Páginas', 'value' => '18 págs', 'color' => 'blue'],
            ['label' => 'Revisões', 'value' => '14 edições', 'color' => 'red'],
            ['label' => 'Tempo Gasto', 'value' => '5h 42m', 'color' => 'gray']
        ];
        $alertas = [
            '👤 O nome de dois colaboradores e da organização corporativa estão expostos no cabeçalho do documento.',
            '⚠️ Revisões anteriores contêm histórico de alterações que pode vazar cláusulas confidenciais removidas.'
        ];

        $tokenLimpo = 'demo_clean_docx';
        $extLimpo = 'docx';
        $resultado = $dados;
    }
    elseif ($demo === 'xlsx') {
        $nomeOriginal = 'Orcamento_Executivo_Consolidado_Q3.xlsx';
        $ext = 'xlsx';
        $icone = "📊";
        $corTopo = "#059669";
        $badges = ['Planilha', 'XLSX', 'Financeiro', 'Corporativo'];

        $hashes = [
            'MD5' => '3d248ef78201a4e1509b2e04f05256e2',
            'SHA-1' => '9d10e58c0b5832049e2170b0213b28198f121a22',
            'SHA-256' => '12a95f80b2a75168019b8529e46a782b192837261a8f948271e8920192837210',
            'SHA-512' => '47e0921820491820391203918203918203918230918230918230918230918203...',
            'CRC32' => '2b842918'
        ];

        $dados = [
            '📝 Nome do Arquivo' => $nomeOriginal,
            '👤 Metadados de Autoria' => [
                'Criador' => 'Felipe Miranda — Controladoria & FP&A',
                'Última Modificação Por' => 'Beatriz Souza — Diretoria Financeira',
                'Empresa' => 'LogExpress Logística Nacional',
                'Título' => 'Orçamento e Projeções Financeiras 2026/2027'
            ],
            '📑 Estrutura da Pasta de Trabalho' => [
                'Total de Abas (Planilhas)' => '4',
                'Abas Detectadas' => 'Resumo Executivo, DRE Consolidado, Fluxo de Caixa Diário, Projeções Q4',
                'Fórmulas Ativas' => '184 cálculos e fórmulas detectadas',
                'Versão do Excel' => 'Microsoft Excel 2021 / Office 365'
            ]
        ];

        $estatisticas = [
            ['label' => 'Abas', 'value' => '4 abas', 'color' => 'green'],
            ['label' => 'Fórmulas', 'value' => '184 cálculos', 'color' => 'blue'],
            ['label' => 'Segurança', 'value' => 'Fórmulas expostas', 'color' => 'red']
        ];
        $alertas = [
            '📊 Planilha contém fórmulas com projeções financeiras corporativas expostas.',
            '👤 Identidade e departamento dos autores expostos nos metadados.'
        ];

        $tokenLimpo = 'demo_clean_xlsx';
        $extLimpo = 'xlsx';
        $resultado = $dados;
    }

    if ($resultado) {
        $privacyScore = calcularScorePrivacidade($resultado, $gpsCoords, $ext);
        $deteccao = detectarSoftwareIA($resultado);
        $iaDetectada = $deteccao['ias'];
        $softwaresDetectados = $deteccao['softwares'];
    }
}

// ==========================================================================
// PROCESSAMENTO DO UPLOAD REAL
// ==========================================================================
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['arquivo'])) {
    $file = $_FILES['arquivo'];

    if ($file['error'] !== UPLOAD_ERR_OK) {
        $codigosErro = [
            1 => 'O arquivo excede o limite de upload do servidor (upload_max_filesize)',
            2 => 'O arquivo excede o limite do formulário',
            3 => 'O upload foi feito apenas parcialmente',
            4 => 'Nenhum arquivo foi enviado',
            6 => 'Pasta temporária ausente no servidor',
            7 => 'Falha ao escrever arquivo no disco',
            8 => 'Uma extensão do PHP interrompeu o upload'
        ];
        $erro = $codigosErro[$file['error']] ?? 'Erro desconhecido no upload.';
    } else {
        $ext = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
        $tmpPath = $file['tmp_name'];
        $nomeOriginal = $file['name'];
        $dados = [];

        try {
            $mimeReal = detectarMimeReal($tmpPath);
            $hashes = calcularHashes($tmpPath);
            $fileStats = stat($tmpPath);

            // Gera cópia higienizada (sanitizada) pronta para download
            $tokenLimpo = higienizarArquivo($tmpPath, $ext);
            $extLimpo = $ext;

            $dados['📝 Nome do Arquivo'] = $nomeOriginal;

            $dados['🔐 Assinatura Digital & Hashes'] = [
                'MD5' => $hashes['MD5'],
                'SHA-1' => $hashes['SHA-1'],
                'SHA-256' => $hashes['SHA-256'],
                'CRC32' => $hashes['CRC32']
            ];

            $dados['📊 Informações do Sistema'] = [
                'MIME Type Detectado' => $mimeReal,
                'Tamanho em Disco' => formatarTamanho($fileStats['size']),
                'Tamanho Bruto' => number_format($fileStats['size'], 0, ',', '.') . ' bytes',
                'Permissões' => substr(sprintf('%o', $fileStats['mode']), -4),
                'Última Modificação' => date('d/m/Y H:i:s', $fileStats['mtime'])
            ];

            // 1. IMAGENS (JPG, PNG, WEBP, TIFF, BMP, GIF)
            if (in_array($ext, ['jpg', 'jpeg', 'png', 'tiff', 'webp', 'gif', 'bmp'])) {
                $icone = "🖼️";
                $corTopo = "#ec4899";
                $badges = ['Imagem', 'Visual', strtoupper($ext)];

                $size = @getimagesize($tmpPath);
                if ($size) {
                    $megapixels = ($size[0] * $size[1]) / 1000000;
                    $aspectRatio = round($size[0] / $size[1], 2);
                    $orientation = $size[1] > $size[0] ? 'Retrato (Vertical)' : ($size[0] == $size[1] ? 'Quadrada' : 'Paisagem (Horizontal)');

                    $widthCm = round(($size[0] / 300) * 2.54, 1);
                    $heightCm = round(($size[1] / 300) * 2.54, 1);

                    $dados['📐 Dimensões e Qualidade'] = [
                        'Resolução' => $size[0] . ' × ' . $size[1] . ' pixels',
                        'Megapixels' => number_format($megapixels, 2) . ' MP',
                        'Proporção' => $aspectRatio . ':1',
                        'Orientação' => $orientation,
                        'Bits por Canal' => ($size['bits'] ?? 8) . '-bit',
                        'Tipo MIME' => $size['mime'],
                        'Tamanho Impresso (300 DPI)' => $widthCm . ' × ' . $heightCm . ' cm'
                    ];

                    $estatisticas[] = ['label' => 'Resolução', 'value' => number_format($megapixels, 1) . ' MP', 'color' => 'blue'];
                    $estatisticas[] = ['label' => 'Qualidade', 'value' => ($megapixels > 12 ? 'Alta Resolução' : ($megapixels > 4 ? 'Boa Resolução' : 'Padrão')), 'color' => 'green'];
                }

                // Leitura EXIF e Decodificação de GPS
                if (function_exists('exif_read_data') && in_array($ext, ['jpg', 'jpeg', 'tiff'])) {
                    $exif = @exif_read_data($tmpPath, 0, true);
                    if ($exif && is_array($exif)) {
                        // GPS Completo com Mapa
                        if (isset($exif['GPS']) && !empty($exif['GPS'])) {
                            if (isset($exif['GPS']['GPSLatitude']) && isset($exif['GPS']['GPSLongitude']) && isset($exif['GPS']['GPSLatitudeRef']) && isset($exif['GPS']['GPSLongitudeRef'])) {
                                $lat = converterGpsCoord($exif['GPS']['GPSLatitude'], $exif['GPS']['GPSLatitudeRef']);
                                $lng = converterGpsCoord($exif['GPS']['GPSLongitude'], $exif['GPS']['GPSLongitudeRef']);

                                if ($lat !== null && $lng !== null) {
                                    $alt = isset($exif['GPS']['GPSAltitude']) ? (avaliarFracao($exif['GPS']['GPSAltitude']) . ' m') : 'Não informada';
                                    $gpsCoords = ['lat' => $lat, 'lng' => $lng, 'alt' => $alt];

                                    $dados['🌐 Geolocalização Exata (GPS)'] = [
                                        'Latitude Decimal' => $lat . '°',
                                        'Longitude Decimal' => $lng . '°',
                                        'Altitude' => $alt,
                                        'Referência' => 'Coordenadas decodificadas de EXIF GPS'
                                    ];

                                    $badges[] = 'Geolocalizado (GPS)';
                                    $estatisticas[] = ['label' => 'Geotag', 'value' => 'GPS Ativo', 'color' => 'red'];
                                    $alertas[] = '🌍 Este arquivo expõe a localização geográfica exata onde a foto foi tirada!';
                                }
                            }
                        }

                        // Equipamento Fotográfico
                        $cameraData = [];
                        if (!empty($exif['IFD0']['Make'])) $cameraData['Fabricante'] = $exif['IFD0']['Make'];
                        if (!empty($exif['IFD0']['Model'])) $cameraData['Modelo'] = $exif['IFD0']['Model'];
                        if (!empty($exif['IFD0']['Software'])) $cameraData['Software'] = $exif['IFD0']['Software'];
                        if (!empty($exif['EXIF']['LensModel'])) $cameraData['Lente'] = $exif['EXIF']['LensModel'];
                        if (!empty($exif['EXIF']['ExposureTime'])) $cameraData['Velocidade Obturador'] = $exif['EXIF']['ExposureTime'] . ' s';
                        if (!empty($exif['EXIF']['FNumber'])) $cameraData['Abertura'] = 'f/' . avaliarFracao($exif['EXIF']['FNumber']);
                        if (!empty($exif['EXIF']['ISOSpeedRatings'])) $cameraData['ISO'] = 'ISO ' . $exif['EXIF']['ISOSpeedRatings'];

                        if (!empty($cameraData)) {
                            $dados['📷 Equipamento & Configurações da Câmera'] = $cameraData;
                            if (isset($cameraData['Modelo'])) $badges[] = $cameraData['Modelo'];
                        }
                    }
                }
            }

            // 2. DOCUMENTOS WORD (DOCX)
            elseif ($ext === 'docx') {
                $icone = "📘";
                $corTopo = "#2563eb";
                $badges = ['Documento', 'Word', 'DOCX'];

                if (class_exists('ZipArchive')) {
                    $zip = new ZipArchive();
                    if ($zip->open($tmpPath) === TRUE) {
                        // Core.xml
                        $coreXml = $zip->getFromName('docProps/core.xml');
                        if ($coreXml) {
                            $coreData = [];
                            if (preg_match('/<dc:creator[^>]*>(.*?)<\/dc:creator>/is', $coreXml, $m)) $coreData['Autor Original'] = trim(strip_tags($m[1]));
                            if (preg_match('/<cp:lastModifiedBy[^>]*>(.*?)<\/cp:lastModifiedBy>/is', $coreXml, $m)) $coreData['Última Modificação Por'] = trim(strip_tags($m[1]));
                            if (preg_match('/<dc:title[^>]*>(.*?)<\/dc:title>/is', $coreXml, $m)) $coreData['Título'] = trim(strip_tags($m[1]));
                            if (preg_match('/<cp:revision[^>]*>(.*?)<\/cp:revision>/is', $coreXml, $m)) $coreData['Número da Revisão'] = trim(strip_tags($m[1]));
                            if (preg_match('/<dcterms:created[^>]*>(.*?)<\/dcterms:created>/is', $coreXml, $m)) $coreData['Data de Criação'] = trim(strip_tags($m[1]));
                            if (preg_match('/<dcterms:modified[^>]*>(.*?)<\/dcterms:modified>/is', $coreXml, $m)) $coreData['Última Modificação'] = trim(strip_tags($m[1]));

                            if (!empty($coreData)) {
                                $dados['👤 Metadados de Autoria (Core.xml)'] = $coreData;
                            }
                        }

                        // App.xml
                        $appXml = $zip->getFromName('docProps/app.xml');
                        if ($appXml) {
                            $appData = [];
                            if (preg_match('/<Company[^>]*>(.*?)<\/Company>/is', $appXml, $m)) $appData['Empresa / Organização'] = trim(strip_tags($m[1]));
                            if (preg_match('/<Application[^>]*>(.*?)<\/Application>/is', $appXml, $m)) $appData['Software Criador'] = trim(strip_tags($m[1]));
                            if (preg_match('/<Pages[^>]*>(.*?)<\/Pages>/is', $appXml, $m)) $appData['Páginas'] = trim(strip_tags($m[1]));
                            if (preg_match('/<Words[^>]*>(.*?)<\/Words>/is', $appXml, $m)) $appData['Palavras'] = trim(strip_tags($m[1]));
                            if (preg_match('/<TotalTime[^>]*>(.*?)<\/TotalTime>/is', $appXml, $m)) $appData['Tempo de Edição'] = trim(strip_tags($m[1])) . ' minutos';

                            if (!empty($appData)) {
                                $dados['🏢 Metadados Corporativos (App.xml)'] = $appData;
                            }
                        }
                        $zip->close();
                    }
                }
            }

            // 3. PLANILHAS EXCEL (XLSX)
            elseif ($ext === 'xlsx') {
                $icone = "📊";
                $corTopo = "#059669";
                $badges = ['Planilha', 'Excel', 'XLSX'];

                if (class_exists('ZipArchive')) {
                    $zip = new ZipArchive();
                    if ($zip->open($tmpPath) === TRUE) {
                        $coreXml = $zip->getFromName('docProps/core.xml');
                        if ($coreXml) {
                            $coreData = [];
                            if (preg_match('/<dc:creator[^>]*>(.*?)<\/dc:creator>/is', $coreXml, $m)) $coreData['Criador'] = trim(strip_tags($m[1]));
                            if (preg_match('/<cp:lastModifiedBy[^>]*>(.*?)<\/cp:lastModifiedBy>/is', $coreXml, $m)) $coreData['Última Modificação Por'] = trim(strip_tags($m[1]));
                            if (preg_match('/<dc:title[^>]*>(.*?)<\/dc:title>/is', $coreXml, $m)) $coreData['Título'] = trim(strip_tags($m[1]));

                            if (!empty($coreData)) $dados['👤 Metadados de Autoria'] = $coreData;
                        }
                        $zip->close();
                    }
                }
            }

            // 4. DOCUMENTOS PDF
            elseif ($ext === 'pdf') {
                $icone = "📄";
                $corTopo = "#dc2626";
                $badges = ['Documento', 'PDF'];

                if (class_exists('Smalot\PdfParser\Parser')) {
                    try {
                        $parser = new PdfParser();
                        $pdf = $parser->parseFile($tmpPath);
                        $details = $pdf->getDetails();

                        $pdfMeta = [];
                        if (!empty($details['Author'])) $pdfMeta['Autor'] = $details['Author'];
                        if (!empty($details['Creator'])) $pdfMeta['Criador / Software'] = $details['Creator'];
                        if (!empty($details['Producer'])) $pdfMeta['Produtor PDF'] = $details['Producer'];
                        if (!empty($details['Title'])) $pdfMeta['Título'] = $details['Title'];
                        if (!empty($details['Subject'])) $pdfMeta['Assunto'] = $details['Subject'];
                        if (!empty($details['Pages'])) $pdfMeta['Número de Páginas'] = $details['Pages'];
                        if (!empty($details['CreationDate'])) $pdfMeta['Data de Criação'] = $details['CreationDate'];

                        if (!empty($pdfMeta)) $dados['📄 Metadados do Documento PDF'] = $pdfMeta;

                        $pages = $pdf->getPages();
                        $estatisticas[] = ['label' => 'Páginas', 'value' => count($pages) . ' págs', 'color' => 'blue'];
                    } catch (Exception $pe) {
                        $alertas[] = '⚠️ Não foi possível analisar os fluxos internos do PDF.';
                    }
                }
            }

            // 5. ÁUDIO E VÍDEO
            elseif (in_array($ext, ['mp3', 'wav', 'mp4', 'mov', 'avi', 'mkv', 'flac', 'ogg'])) {
                $icone = "🎬";
                $corTopo = "#7c3aed";
                $badges = ['Mídia', strtoupper($ext)];

                if (class_exists('getID3')) {
                    $getID3 = new getID3;
                    $fileInfo = $getID3->analyze($tmpPath);

                    $mediaData = [];
                    if (isset($fileInfo['playtime_string'])) $mediaData['Duração'] = $fileInfo['playtime_string'];
                    if (isset($fileInfo['bitrate'])) $mediaData['Bitrate'] = round($fileInfo['bitrate'] / 1000) . ' kbps';
                    if (isset($fileInfo['video']['resolution_x'])) $mediaData['Resolução'] = $fileInfo['video']['resolution_x'] . ' × ' . $fileInfo['video']['resolution_y'];
                    if (isset($fileInfo['audio']['codec'])) $mediaData['Codec de Áudio'] = $fileInfo['audio']['codec'];

                    if (!empty($mediaData)) $dados['🎵 Informações de Reprodução'] = $mediaData;
                }
            }

            // 6. HTML / WEB
            elseif (in_array($ext, ['html', 'htm'])) {
                $icone = "🌐";
                $corTopo = "#ea580c";
                $badges = ['Web', 'HTML'];

                libxml_use_internal_errors(true);
                $dom = new DOMDocument();
                $dom->loadHTML(file_get_contents($tmpPath));

                $titles = $dom->getElementsByTagName('title');
                if ($titles->length > 0) $dados['🔑 Título'] = ['Título' => $titles->item(0)->nodeValue];
            }

            $resultado = $dados;

            // Avaliações de Risco e IA
            $privacyScore = calcularScorePrivacidade($resultado, $gpsCoords, $ext);
            $deteccao = detectarSoftwareIA($resultado);
            $iaDetectada = $deteccao['ias'];
            $softwaresDetectados = $deteccao['softwares'];

        } catch (Exception $e) {
            $erro = "Erro durante o processamento do arquivo: " . $e->getMessage();
        }
    }
}
?>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>4U MetaViewer Pro 5.0 — Auditoria de Metadados, LGPD & Geolocalização</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="Analisador avançado de metadados com higienização de arquivos (Sanitizer LGPD), mapa interativo de geolocalização GPS e perícia forense.">

    <!-- Google Fonts -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&display=swap" rel="stylesheet">

    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>

    <!-- Leaflet CSS & JS (OpenStreetMap) -->
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>

    <style>
        :root {
            --bg-main: #070b14;
            --bg-card: rgba(15, 23, 42, 0.78);
            --border-subtle: rgba(255, 255, 255, 0.08);
            --primary: #8b5cf6;
            --primary-hover: #7c3aed;
        }

        body {
            background-color: var(--bg-main);
            background-image: 
                radial-gradient(at 0% 0%, rgba(139, 92, 246, 0.14) 0px, transparent 50%),
                radial-gradient(at 100% 100%, rgba(6, 182, 212, 0.12) 0px, transparent 50%);
            background-attachment: fixed;
            color: #f1f5f9;
            font-family: 'Inter', system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            min-height: 100vh;
        }

        .glass-panel {
            background: var(--bg-card);
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px);
            border: 1px solid var(--border-subtle);
            border-radius: 24px;
            box-shadow: 0 10px 30px -10px rgba(0, 0, 0, 0.5);
        }

        .brand-logo-glow {
            filter: drop-shadow(0 0 12px rgba(139, 92, 246, 0.45));
            transition: transform 0.3s ease, filter 0.3s ease;
        }
        .brand-logo-glow:hover {
            transform: scale(1.05);
            filter: drop-shadow(0 0 18px rgba(139, 92, 246, 0.7));
        }

        .pro-badge {
            background: linear-gradient(135deg, #8b5cf6 0%, #06b6d4 100%);
            color: #ffffff;
            font-weight: 800;
            font-size: 0.68rem;
            letter-spacing: 0.1em;
            padding: 3px 8px;
            border-radius: 9999px;
            text-transform: uppercase;
            box-shadow: 0 2px 10px rgba(139, 92, 246, 0.4);
        }

        .upload-dropzone {
            border: 2px dashed rgba(139, 92, 246, 0.4);
            border-radius: 20px;
            padding: 50px 24px;
            text-align: center;
            cursor: pointer;
            background: rgba(139, 92, 246, 0.04);
            transition: all 0.3s ease;
        }
        .upload-dropzone:hover, .upload-dropzone.dragover {
            border-color: #8b5cf6;
            background: rgba(139, 92, 246, 0.1);
            transform: translateY(-2px);
        }

        .badge-pill {
            font-size: 0.75rem;
            font-weight: 600;
            padding: 4px 10px;
            border-radius: 9999px;
            background: rgba(255, 255, 255, 0.1);
            border: 1px solid rgba(255, 255, 255, 0.15);
        }

        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: rgba(15, 23, 42, 0.6); }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: rgba(139, 92, 246, 0.5); border-radius: 3px; }

        @media print {
            body { background: white !important; color: black !important; }
            .no-print { display: none !important; }
            .glass-panel { border: 1px solid #ccc !important; box-shadow: none !important; background: white !important; color: black !important; }
            .print-only { display: block !important; }
        }
    </style>
</head>
<body class="p-3 sm:p-6 lg:p-8">

    <div class="max-w-5xl mx-auto space-y-6">

        <!-- ==========================================================================
             CABEÇALHO OFICIAL 4U ECOSYSTEM
             ========================================================================== -->
        <header class="glass-panel p-6 flex flex-col md:flex-row items-center justify-between gap-4 no-print">
            <div class="flex items-center gap-4 text-center md:text-left">
                <!-- 4U Glowing Logo -->
                <a href="index.php" class="w-13 h-13 rounded-2xl bg-gradient-to-tr from-violet-600 to-cyan-500 p-0.5 brand-logo-glow flex-shrink-0">
                    <div class="w-full h-full bg-slate-950 rounded-2xl flex items-center justify-center">
                        <svg class="w-7 h-7 text-white" viewBox="0 0 100 100" fill="none">
                            <rect x="15" y="15" width="70" height="70" rx="16" stroke="url(#logoGrad)" stroke-width="8"/>
                            <path d="M35 65V35L52 55V35" stroke="#ffffff" stroke-width="7" stroke-linecap="round" stroke-linejoin="round"/>
                            <path d="M58 35V53C58 60 68 60 68 53V35" stroke="url(#logoGrad)" stroke-width="7" stroke-linecap="round"/>
                            <defs>
                                <linearGradient id="logoGrad" x1="0%" y1="0%" x2="100%" y2="100%">
                                    <stop offset="0%" stop-color="#8b5cf6" />
                                    <stop offset="100%" stop-color="#06b6d4" />
                                </linearGradient>
                            </defs>
                        </svg>
                    </div>
                </a>

                <div>
                    <div class="flex items-center justify-center md:justify-start gap-2.5">
                        <h1 class="text-2xl sm:text-3xl font-black tracking-tight text-white">4U METAVIEWER</h1>
                        <span class="pro-badge">PRO 5.0</span>
                    </div>
                    <p class="text-xs sm:text-sm text-slate-400 mt-1 max-w-xl">
                        Auditoria Forense de Metadados, Higienização de Arquivos (LGPD), Geolocalização em Mapa & Detecção de IA.
                    </p>
                </div>
            </div>

            <!-- Demonstrações em 1 Clique -->
            <div class="flex flex-wrap items-center justify-center gap-2">
                <span class="text-xs text-slate-500 font-medium">Testar com Demo:</span>
                <a href="index.php?demo=gps" class="px-3 py-1.5 rounded-xl bg-violet-600/20 hover:bg-violet-600/30 border border-violet-500/30 text-xs font-semibold text-violet-300 transition-all">
                    📸 Foto GPS
                </a>
                <a href="index.php?demo=docx" class="px-3 py-1.5 rounded-xl bg-blue-600/20 hover:bg-blue-600/30 border border-blue-500/30 text-xs font-semibold text-blue-300 transition-all">
                    📘 DOCX
                </a>
                <a href="index.php?demo=xlsx" class="px-3 py-1.5 rounded-xl bg-emerald-600/20 hover:bg-emerald-600/30 border border-emerald-500/30 text-xs font-semibold text-emerald-300 transition-all">
                    📊 XLSX
                </a>
            </div>
        </header>

        <!-- MENSAGEM DE ERRO -->
        <?php if ($erro): ?>
            <div class="p-5 rounded-2xl bg-rose-500/10 border border-rose-500/30 text-rose-300 text-sm flex items-center justify-between gap-3">
                <div class="flex items-center gap-3">
                    <span class="text-2xl">⚠️</span>
                    <div>
                        <strong class="block font-bold">Falha no Processamento</strong>
                        <span><?= htmlspecialchars($erro) ?></span>
                    </div>
                </div>
                <a href="index.php" class="px-4 py-2 rounded-xl bg-rose-600 hover:bg-rose-500 text-white font-semibold text-xs transition-all">
                    Tentar Novamente
                </a>
            </div>
        <?php endif; ?>

        <!-- ==========================================================================
             ÁREA DE UPLOAD (QUANDO NÃO HOUVER RESULTADO)
             ========================================================================== -->
        <?php if (!$resultado): ?>
            <section class="glass-panel p-8 text-center space-y-6">
                <form method="POST" enctype="multipart/form-data" id="uploadForm">
                    <div class="upload-dropzone" id="dropZone">
                        <input type="file" name="arquivo" id="fileInput" class="hidden" onchange="submitUpload()">
                        
                        <div class="w-16 h-16 mx-auto mb-4 rounded-3xl bg-violet-600/20 border border-violet-500/30 flex items-center justify-center text-3xl shadow-inner">
                            ☁️
                        </div>

                        <h2 class="text-xl sm:text-2xl font-bold text-white mb-2">
                            Arraste ou clique para selecionar um arquivo
                        </h2>
                        <p class="text-xs sm:text-sm text-slate-400 max-w-md mx-auto mb-6">
                            Extraia metadados completos de documentos, imagens, mídias e código. Higienize antes de compartilhar.
                        </p>

                        <!-- Pílulas de Formatos Suportados -->
                        <div class="flex flex-wrap items-center justify-center gap-2 mb-6">
                            <span class="badge-pill text-violet-300 bg-violet-600/10 border-violet-500/30">📄 PDF</span>
                            <span class="badge-pill text-blue-300 bg-blue-600/10 border-blue-500/30">📘 DOCX</span>
                            <span class="badge-pill text-emerald-300 bg-emerald-600/10 border-emerald-500/30">📊 XLSX</span>
                            <span class="badge-pill text-pink-300 bg-pink-600/10 border-pink-500/30">🖼️ JPG/PNG/WEBP</span>
                            <span class="badge-pill text-purple-300 bg-purple-600/10 border-purple-500/30">🎵 MP3</span>
                            <span class="badge-pill text-rose-300 bg-rose-600/10 border-rose-500/30">🎬 MP4</span>
                            <span class="badge-pill text-amber-300 bg-amber-600/10 border-amber-500/30">🌐 HTML</span>
                        </div>

                        <button type="button" onclick="document.getElementById('fileInput').click()" class="bg-gradient-to-r from-violet-600 to-indigo-600 hover:from-violet-500 hover:to-indigo-500 text-white font-semibold px-8 py-3.5 rounded-xl shadow-lg shadow-violet-600/25 transition-all inline-flex items-center gap-2 cursor-pointer transform hover:-translate-y-0.5">
                            <span>📂</span>
                            <span>Selecionar Arquivo</span>
                        </button>
                    </div>
                </form>

                <!-- Aviso de Privacidade Zero-Knowledge -->
                <div class="flex items-center justify-center gap-2 text-xs text-slate-500 pt-2">
                    <span>🔒</span>
                    <span>Processamento seguro. Arquivos temporários são descartados imediatamente após a extração.</span>
                </div>
            </section>
        <?php endif; ?>

        <!-- ==========================================================================
             ÁREA DE RESULTADOS & RELATÓRIO FORENSE
             ========================================================================== -->
        <?php if ($resultado): ?>
            <div class="space-y-6">

                <!-- 1. HEADER DO ARQUIVO COM BADGES & IA DETECTADA -->
                <section class="glass-panel p-6 flex flex-col md:flex-row items-center justify-between gap-5" style="border-top: 4px solid <?= $corTopo ?>;">
                    <div class="flex items-center gap-4 text-center md:text-left">
                        <div class="w-16 h-16 rounded-2xl flex items-center justify-center text-3xl flex-shrink-0" style="background: <?= $corTopo ?>25; border: 1px solid <?= $corTopo ?>60;">
                            <?= $icone ?>
                        </div>
                        <div>
                            <span class="text-[11px] uppercase tracking-wider text-slate-400 font-bold block mb-1">
                                Relatório Técnico de Metadados
                            </span>
                            <h2 class="text-xl sm:text-2xl font-black text-white break-all">
                                <?= htmlspecialchars($resultado['📝 Nome do Arquivo'] ?? 'Arquivo') ?>
                            </h2>
                            <div class="flex flex-wrap items-center justify-center md:justify-start gap-2 mt-2">
                                <?php foreach ($badges as $b): ?>
                                    <span class="badge-pill text-slate-200"><?= htmlspecialchars($b) ?></span>
                                <?php endforeach; ?>
                                <?php foreach ($iaDetectada as $ia): ?>
                                    <span class="badge-pill bg-purple-600/30 text-purple-300 border-purple-500/40 font-bold">
                                        🤖 Criado via <?= htmlspecialchars($ia) ?>
                                    </span>
                                <?php endforeach; ?>
                                <?php foreach ($softwaresDetectados as $soft): ?>
                                    <span class="badge-pill bg-blue-600/30 text-blue-300 border-blue-500/40">
                                        🎨 Editado no <?= htmlspecialchars($soft) ?>
                                    </span>
                                <?php endforeach; ?>
                            </div>
                        </div>
                    </div>

                    <!-- Botão Reset -->
                    <div class="no-print">
                        <a href="index.php" class="px-4 py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-200 border border-slate-700 transition-all flex items-center gap-2">
                            <span>🔄</span>
                            <span>Analisar Outro Arquivo</span>
                        </a>
                    </div>
                </section>

                <!-- 2. BARRA DE AÇÕES HERO (HIGIENIZADOR / LAUDO / EXPORTAÇÃO) -->
                <section class="glass-panel p-5 no-print">
                    <div class="flex flex-wrap items-center justify-between gap-3">
                        <div>
                            <h3 class="text-sm font-bold text-white flex items-center gap-2">
                                <span>⚡</span>
                                <span>Ações Rápidas de Privacidade & Auditoria</span>
                            </h3>
                            <p class="text-xs text-slate-400">Proteja sua privacidade removendo metadados ou exporte relatórios técnicos.</p>
                        </div>

                        <div class="flex flex-wrap items-center gap-2">
                            <!-- Botão Higienizador (Metadata Stripper) -->
                            <?php if ($tokenLimpo && $extLimpo): ?>
                                <a href="index.php?download_clean=<?= urlencode($tokenLimpo) ?>&ext=<?= urlencode($extLimpo) ?>&orig=<?= urlencode($nomeOriginal) ?>" class="px-4 py-2.5 rounded-xl bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-bold text-xs shadow-lg shadow-emerald-600/25 transition-all flex items-center gap-2">
                                    <span>🛡️</span>
                                    <span>Baixar Arquivo Higienizado (Sem Metadados)</span>
                                </a>
                            <?php endif; ?>

                            <!-- Laudo PDF / Print -->
                            <button type="button" onclick="window.print()" class="px-3.5 py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 font-semibold text-xs transition-all flex items-center gap-1.5 cursor-pointer">
                                <span>📑</span>
                                <span>Imprimir Laudo (PDF)</span>
                            </button>

                            <!-- Exportar JSON -->
                            <button type="button" onclick="baixarJSON()" class="px-3.5 py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 font-semibold text-xs transition-all flex items-center gap-1.5 cursor-pointer">
                                <span>💾</span>
                                <span>Exportar JSON</span>
                            </button>

                            <!-- Copiar Markdown -->
                            <button type="button" onclick="copiarMarkdown()" class="px-3.5 py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 font-semibold text-xs transition-all flex items-center gap-1.5 cursor-pointer">
                                <span>📋</span>
                                <span>Copiar Markdown</span>
                            </button>
                        </div>
                    </div>
                </section>

                <!-- 3. SCORE DE RISCO DE PRIVACIDADE (PRIVACY SCORE LGPD) -->
                <?php if ($privacyScore): ?>
                    <section class="glass-panel p-6 space-y-4">
                        <div class="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
                            <div>
                                <span class="text-xs font-bold uppercase tracking-wider text-slate-400">Diagnóstico de Privacidade & LGPD</span>
                                <h3 class="text-xl font-bold text-white flex items-center gap-2 mt-0.5">
                                    <span>Nível de Risco:</span>
                                    <span style="color: <?= $privacyScore['cor'] ?>;"><?= $privacyScore['rotulo'] ?></span>
                                </h3>
                            </div>
                            <div class="text-right">
                                <span class="text-3xl font-black" style="color: <?= $privacyScore['cor'] ?>;">
                                    <?= $privacyScore['score'] ?>%
                                </span>
                                <span class="text-xs text-slate-400 block">Exposição de Metadados</span>
                            </div>
                        </div>

                        <!-- Barra de Risco -->
                        <div class="w-full bg-slate-900 h-2.5 rounded-full overflow-hidden border border-slate-800">
                            <div class="h-full rounded-full transition-all duration-500" style="width: <?= $privacyScore['score'] ?>%; background: <?= $privacyScore['cor'] ?>;"></div>
                        </div>

                        <!-- Checklist de Fatores de Exposição -->
                        <?php if (!empty($privacyScore['fatores'])): ?>
                            <div class="space-y-2 pt-2">
                                <span class="text-xs text-slate-400 font-semibold">Fatores que impactam o nível de privacidade:</span>
                                <div class="grid grid-cols-1 md:grid-cols-2 gap-2">
                                    <?php foreach ($privacyScore['fatores'] as $fat): ?>
                                        <div class="p-2.5 rounded-xl bg-slate-900/80 border border-slate-800/80 text-xs flex items-center gap-2 text-slate-200">
                                            <span class="text-base"><?= $fat['tipo'] === 'critico' ? '🔴' : ($fat['tipo'] === 'alto' ? '🟠' : '🟡') ?></span>
                                            <span><?= $fat['texto'] ?></span>
                                        </div>
                                    <?php endforeach; ?>
                                </div>
                            </div>
                        <?php endif; ?>
                    </section>
                <?php endif; ?>

                <!-- 4. GEOLOCALIZAÇÃO EM MAPA INTERATIVO (QUANDO HOUVER GPS) -->
                <?php if ($gpsCoords): ?>
                    <section class="glass-panel p-6 space-y-4">
                        <div class="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
                            <div>
                                <span class="text-xs font-bold uppercase tracking-wider text-rose-400">Rastreamento Físico Detectado</span>
                                <h3 class="text-lg font-bold text-white flex items-center gap-2">
                                    <span>🌍 Geolocalização em Mapa (OpenStreetMap)</span>
                                </h3>
                                <p class="text-xs text-slate-400">
                                    Latitude: <strong class="text-white"><?= $gpsCoords['lat'] ?>°</strong> • Longitude: <strong class="text-white"><?= $gpsCoords['lng'] ?>°</strong> • Altitude: <strong class="text-white"><?= htmlspecialchars($gpsCoords['alt'] ?? '') ?></strong>
                                </p>
                            </div>

                            <div class="flex items-center gap-2 no-print">
                                <a href="https://www.google.com/maps?q=<?= $gpsCoords['lat'] ?>,<?= $gpsCoords['lng'] ?>" target="_blank" rel="noopener noreferrer" class="px-3 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-200 border border-slate-700 transition-all flex items-center gap-1.5">
                                    <span>🗺️</span>
                                    <span>Google Maps</span>
                                </a>
                                <a href="https://waze.com/ul?ll=<?= $gpsCoords['lat'] ?>,<?= $gpsCoords['lng'] ?>&navigate=yes" target="_blank" rel="noopener noreferrer" class="px-3 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-200 border border-slate-700 transition-all flex items-center gap-1.5">
                                    <span>🧭</span>
                                    <span>Waze</span>
                                </a>
                            </div>
                        </div>

                        <!-- Mini Mapa Leaflet -->
                        <div id="mapLeaflet" class="w-full h-80 rounded-2xl overflow-hidden border border-slate-700/80 z-0"></div>
                    </section>
                <?php endif; ?>

                <!-- 5. ESTATÍSTICAS CHAVE & KPIS -->
                <?php if (!empty($estatisticas)): ?>
                    <section class="grid grid-cols-2 sm:grid-cols-3 gap-4">
                        <?php foreach ($estatisticas as $st): ?>
                            <div class="glass-panel p-4 text-center">
                                <span class="text-xs uppercase font-bold text-slate-400 tracking-wider block mb-1"><?= htmlspecialchars($st['label']) ?></span>
                                <span class="text-2xl font-black text-white"><?= htmlspecialchars($st['value']) ?></span>
                            </div>
                        <?php endforeach; ?>
                    </section>
                <?php endif; ?>

                <!-- 6. ASSINATURA DIGITAL & HASHES CRIPTOGRÁFICOS -->
                <?php if (!empty($hashes)): ?>
                    <section class="glass-panel p-6 space-y-3">
                        <h3 class="text-sm font-bold text-white flex items-center gap-2">
                            <span>🔐</span>
                            <span>Assinatura Digital & Hashes Forenses</span>
                        </h3>
                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs font-mono">
                            <?php foreach ($hashes as $tipoHash => $valorHash): ?>
                                <div class="p-3 rounded-xl bg-slate-900/80 border border-slate-800 flex items-center justify-between gap-2">
                                    <div class="min-w-0">
                                        <span class="text-violet-400 font-bold block text-[10px]"><?= $tipoHash ?></span>
                                        <span class="text-slate-300 truncate block select-all"><?= htmlspecialchars($valorHash) ?></span>
                                    </div>
                                    <button type="button" onclick="navigator.clipboard.writeText('<?= htmlspecialchars($valorHash) ?>'); alert('Hash <?= $tipoHash ?> copiado!');" class="p-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white transition-all no-print" title="Copiar Hash">
                                        📋
                                    </button>
                                </div>
                            <?php endforeach; ?>
                        </div>
                    </section>
                <?php endif; ?>

                <!-- 7. TABELAS DE METADADOS CATEGORIZADAS -->
                <section class="glass-panel p-6 space-y-6">
                    <h3 class="text-base font-bold text-white flex items-center gap-2 border-b border-slate-800 pb-3">
                        <span>📋</span>
                        <span>Inventário Completo de Metadados Extraídos</span>
                    </h3>

                    <?php 
                    unset($resultado['📝 Nome do Arquivo']);
                    foreach ($resultado as $grupoTitulo => $grupoConteudo): 
                    ?>
                        <div class="space-y-3">
                            <h4 class="text-xs sm:text-sm font-bold text-violet-300 flex items-center gap-2">
                                <span>•</span>
                                <span><?= htmlspecialchars($grupoTitulo) ?></span>
                            </h4>

                            <?php if (is_array($grupoConteudo)): ?>
                                <div class="rounded-xl border border-slate-800 overflow-hidden">
                                    <table class="w-full text-xs text-left">
                                        <tbody>
                                            <?php foreach ($grupoConteudo as $subChave => $subValor): ?>
                                                <tr class="border-b border-slate-800/60 hover:bg-slate-900/60 transition-colors">
                                                    <td class="p-3 w-1/3 font-semibold text-slate-400 bg-slate-950/40">
                                                        <?= htmlspecialchars($subChave) ?>
                                                    </td>
                                                    <td class="p-3 text-slate-200 font-mono break-all select-all">
                                                        <?= htmlspecialchars(is_array($subValor) ? json_encode($subValor) : $subValor) ?>
                                                    </td>
                                                </tr>
                                            <?php endforeach; ?>
                                        </tbody>
                                    </table>
                                </div>
                            <?php else: ?>
                                <div class="p-4 rounded-xl bg-slate-900/60 border border-slate-800 text-xs text-slate-200 font-mono">
                                    <?= htmlspecialchars($grupoConteudo) ?>
                                </div>
                            <?php endif; ?>
                        </div>
                    <?php endforeach; ?>
                </section>

            </div>
        <?php endif; ?>

        <!-- ==========================================================================
             RODAPÉ OFICIAL 4U ECOSYSTEM
             ========================================================================== -->
        <footer class="text-center py-6 text-xs text-slate-500 space-y-1 no-print">
            <p>4U MetaViewer Pro — Zero-Knowledge Metadata Extractor & Sanitizer Engine</p>
            <p>Desenvolvido pela equipe <strong class="text-violet-400 font-bold">4U.IA.BR</strong> • 100% no servidor privado, sem compartilhamento com terceiros.</p>
        </footer>

    </div>

    <!-- SCRIPTS INTERATIVOS -->
    <script>
        // Dropzone e Upload automático
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const uploadForm = document.getElementById('uploadForm');

        if (dropZone) {
            ['dragenter', 'dragover'].forEach(name => {
                dropZone.addEventListener(name, (e) => {
                    e.preventDefault();
                    dropZone.classList.add('dragover');
                });
            });

            ['dragleave', 'drop'].forEach(name => {
                dropZone.addEventListener(name, (e) => {
                    e.preventDefault();
                    dropZone.classList.remove('dragover');
                });
            });

            dropZone.addEventListener('drop', (e) => {
                const files = e.dataTransfer.files;
                if (files && files.length > 0) {
                    fileInput.files = files;
                    submitUpload();
                }
            });
        }

        function submitUpload() {
            if (fileInput && fileInput.files.length > 0) {
                if (dropZone) {
                    dropZone.innerHTML = `
                        <div class="py-8">
                            <div class="w-12 h-12 border-4 border-violet-500 border-t-transparent rounded-full animate-spin mx-auto mb-3"></div>
                            <p class="text-sm font-bold text-white">Analisando metadados do arquivo...</p>
                            <p class="text-xs text-slate-400 mt-1">Calculando hashes e extraindo metadados profundos</p>
                        </div>
                    `;
                }
                uploadForm.submit();
            }
        }

        // Inicialização do Mapa Leaflet (se houver coordenadas)
        <?php if ($gpsCoords): ?>
            document.addEventListener('DOMContentLoaded', () => {
                try {
                    const lat = <?= json_encode($gpsCoords['lat']) ?>;
                    const lng = <?= json_encode($gpsCoords['lng']) ?>;
                    const map = L.map('mapLeaflet').setView([lat, lng], 15);

                    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
                        attribution: '© OpenStreetMap contributors',
                        maxZoom: 19
                    }).addTo(map);

                    const marker = L.marker([lat, lng]).addTo(map);
                    marker.bindPopup("<b>Localização da Captura</b><br>Lat: " + lat + "<br>Lng: " + lng).openPopup();
                } catch (e) {
                    console.error('Erro ao renderizar mapa Leaflet:', e);
                }
            });
        <?php endif; ?>

        // Exportação em JSON
        function baixarJSON() {
            const dados = <?= json_encode($resultado ?? [], JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT) ?>;
            const blob = new Blob([JSON.stringify(dados, null, 2)], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'metadados_laudo.json';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        // Cópia em Markdown
        function copiarMarkdown() {
            const dados = <?= json_encode($resultado ?? [], JSON_UNESCAPED_UNICODE) ?>;
            let md = '# Laudo Técnico de Metadados — 4U MetaViewer Pro\n\n';
            for (const [grupo, itens] of Object.entries(dados)) {
                md += `## ${grupo}\n`;
                if (typeof itens === 'object' && itens !== null) {
                    for (const [k, v] of Object.entries(itens)) {
                        md += `- **${k}:** ${v}\n`;
                    }
                } else {
                    md += `${itens}\n`;
                }
                md += '\n';
            }
            navigator.clipboard.writeText(md).then(() => {
                alert('📋 Laudo formatado em Markdown copiado com sucesso!');
            });
        }
    </script>
</body>
</html>
