<?php
// Configurações
$arquivoZip = 'site.zip'; // O nome do arquivo que você vai subir
$caminhoDestino = __DIR__; // __DIR__ significa a pasta atual onde este script está

$mensagem = "";
$status = "aguardando"; // aguardando, sucesso, erro

// Verifica se o usuário clicou no botão
if (isset($_POST['acao']) && $_POST['acao'] == 'descompactar') {
    
    // Verifica se a extensão ZIP está ativa no servidor
    if (!class_exists('ZipArchive')) {
        $mensagem = "Erro Crítico: A extensão PHP 'ZipArchive' não está ativada no servidor.";
        $status = "erro";
    } elseif (!file_exists($arquivoZip)) {
        $mensagem = "Erro: O arquivo <strong>$arquivoZip</strong> não foi encontrado nesta pasta.";
        $status = "erro";
    } else {
        $zip = new ZipArchive;
        if ($zip->open($arquivoZip) === TRUE) {
            // Tenta extrair
            if($zip->extractTo($caminhoDestino)) {
                $zip->close();
                $mensagem = "Sucesso! Arquivos extraídos corretamente.";
                $status = "sucesso";
                
                // Opcional: Deletar o ZIP após extrair para limpar o servidor
                // unlink($arquivoZip); 
            } else {
                $mensagem = "Erro: Falha ao escrever os arquivos. Verifique as permissões da pasta.";
                $status = "erro";
            }
        } else {
            $mensagem = "Erro: Não foi possível abrir o arquivo $arquivoZip. Ele pode estar corrompido.";
            $status = "erro";
        }
    }
}
?>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Descompactador de Site</title>
    <style>
        body { font-family: sans-serif; background: #f4f4f4; display: flex; justify-content: center; align-items: center; height: 100vh; }
        .card { background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); text-align: center; width: 400px; }
        h1 { margin-top: 0; color: #333; }
        .btn { background: #28a745; color: white; border: none; padding: 15px 30px; font-size: 16px; border-radius: 5px; cursor: pointer; width: 100%; transition: 0.3s; }
        .btn:hover { background: #218838; }
        .msg { margin-top: 20px; padding: 15px; border-radius: 5px; }
        .sucesso { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .erro { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .info { font-size: 0.9em; color: #666; margin-bottom: 20px; }
    </style>
</head>
<body>

    <div class="card">
        <h1>📦 Instalador</h1>
        
        <p class="info">Certifique-se de que o arquivo <strong><?php echo $arquivoZip; ?></strong> está na mesma pasta que este script.</p>

        <?php if ($status == 'sucesso'): ?>
            <div class="msg sucesso">
                <?php echo $mensagem; ?>
                <p><a href="index.php">Ir para o App Principal</a></p>
            </div>
            <p style="font-size: 0.8em; margin-top: 15px;">Agora você pode deletar este arquivo descompactador.php</p>
        <?php else: ?>
            
            <?php if ($status == 'erro'): ?>
                <div class="msg erro"><?php echo $mensagem; ?></div>
            <?php endif; ?>

            <form method="POST">
                <input type="hidden" name="acao" value="descompactar">
                <button type="submit" class="btn">Descompactar Agora</button>
            </form>

        <?php endif; ?>
    </div>

</body>
</html>