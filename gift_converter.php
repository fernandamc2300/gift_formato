<?php
// ============================================================
//  GIFT CONVERTER — detecta respuesta correcta con "C" al final
// ============================================================

if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
    $phpwordAvailable = true;
} else {
    $phpwordAvailable = false;
}

// ── Entidades HTML ──────────────────────────────────────────
function toHtmlEntities(string $text): string {
    $map = [
        '¿' => '&iquest;',
        '?' => '&quest;',
        'á' => '&aacute;',
        'é' => '&eacute;',
        'í' => '&iacute;',
        'ó' => '&oacute;',
        'ú' => '&uacute;',
        'Á' => '&Aacute;',
        'É' => '&Eacute;',
        'Í' => '&Iacute;',
        'Ó' => '&Oacute;',
        'Ú' => '&Uacute;',
        'ñ' => '&ntilde;',
        'Ñ' => '&Ntilde;',
        '°' => '&deg;',
    ];
    return strtr($text, $map);
}

// ── Detectar si una opción tiene la C de correcta al final ──
// Acepta: "Texto C", "Texto C.", "Texto  C", "Texto (C)"
function isCorrect(string &$text): bool {
    // Patrones: espacio(s) + C sola al final, opcionalmente con paréntesis o punto
    if (preg_match('/\s+C\.?\s*$/u', $text) ||
        preg_match('/\s+\(C\)\s*$/u', $text)) {
        // Limpiar la C del texto
        $text = trim(preg_replace('/\s+[\(]?C[\)]?\.?\s*$/u', '', $text));
        return true;
    }
    return false;
}

// ── Extraer texto línea a línea de .docx ───────────────────
function extractLinesFromDocx(string $filepath): array {
    if (!class_exists('\PhpOffice\PhpWord\IOFactory')) {
        throw new Exception('PhpWord no disponible.');
    }

    $phpWord = \PhpOffice\PhpWord\IOFactory::load($filepath);
    $lines   = [];

    foreach ($phpWord->getSections() as $section) {
        foreach ($section->getElements() as $element) {
            $text = extractElementText($element);
            if (trim($text) !== '') {
                $lines[] = trim($text);
            }
        }
    }

    return $lines;
}

function extractElementText($element): string {
    $text = '';
    if ($element instanceof \PhpOffice\PhpWord\Element\Text) {
        return $element->getText();
    }
    if (method_exists($element, 'getElements')) {
        foreach ($element->getElements() as $child) {
            $text .= extractElementText($child);
        }
    }
    return $text;
}

// ── Convertir líneas → GIFT ─────────────────────────────────
function convertToGift(array $lines): string {
    $output   = '';
    $question = '';
    $options  = [];

    $flushQuestion = function() use (&$question, &$options, &$output) {
        if ($question === '') return;

        $q = toHtmlEntities(trim($question));
        $output .= "// Start of question: MultipleChoiceQ\n";
        $output .= $q . " {\n";

        foreach ($options as $opt) {
            $text   = toHtmlEntities(trim($opt['text']));
            $marker = $opt['correct'] ? '=' : '~';
            $output .= $marker . $text . "\n";
        }

        $output  .= "}\n\n";
        $question = '';
        $options  = [];
    };

    foreach ($lines as $line) {
        $trimmed = trim($line);
        if ($trimmed === '') continue;

        // Saltar encabezados de sección
        if (preg_match('/^PREGUNTAS TIPO/i', $trimmed)) continue;

        // Número de pregunta: "1.", "21)", etc.
        if (preg_match('/^\d+[\.\)]\s+(.+)/u', $trimmed, $m)) {
            $flushQuestion();
            $question = $m[1];
            continue;
        }

        // Opción con "o " al inicio (formato Word típico)
        if (preg_match('/^o\s+(.+)/u', $trimmed, $m)) {
            $text    = $m[1];
            $correct = isCorrect($text);
            $options[] = ['text' => trim($text), 'correct' => $correct];
            continue;
        }

        // Opción con -, *, •
        if (preg_match('/^[\-\*•]\s+(.+)/u', $trimmed, $m)) {
            $text    = $m[1];
            $correct = isCorrect($text);
            $options[] = ['text' => trim($text), 'correct' => $correct];
            continue;
        }

        // Opción con = o ~ explícito (compatibilidad con formato anterior)
        if (preg_match('/^=\s*(.+)/u', $trimmed, $m)) {
            $options[] = ['text' => trim($m[1]), 'correct' => true];
            continue;
        }
        if (preg_match('/^~\s*(.+)/u', $trimmed, $m)) {
            $options[] = ['text' => trim($m[1]), 'correct' => false];
            continue;
        }

        // Continuación de pregunta (sin opciones aún)
        if ($question !== '' && empty($options)) {
            $question .= ' ' . $trimmed;
        }
    }

    $flushQuestion();
    return trim($output);
}

// ── Manejo del POST ──────────────────────────────────────────
$giftOutput = '';
$error      = '';
$warning    = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    if (isset($_FILES['archivo']) && $_FILES['archivo']['error'] === UPLOAD_ERR_OK) {
        $ext     = strtolower(pathinfo($_FILES['archivo']['name'], PATHINFO_EXTENSION));
        $tmpPath = $_FILES['archivo']['tmp_name'];

        if ($ext === 'txt') {
            $rawLines   = preg_split('/\r\n|\r|\n/', file_get_contents($tmpPath));
            $giftOutput = convertToGift($rawLines);

        } elseif ($ext === 'docx') {
            if (!$phpwordAvailable) {
                $error = 'PhpWord no instalado. Ejecuta: composer require phpoffice/phpword';
            } else {
                try {
                    $lines      = extractLinesFromDocx($tmpPath);
                    $giftOutput = convertToGift($lines);
                } catch (Exception $e) {
                    $error = 'Error al leer el Word: ' . $e->getMessage();
                }
            }

        } elseif (in_array($ext, ['doc', 'dot'])) {
            // Intentar convertir con LibreOffice si está disponible
            $outDir  = sys_get_temp_dir();
            $tmpCopy = $outDir . '/' . basename($tmpPath) . '.' . $ext;
            copy($tmpPath, $tmpCopy);
            $cmd = "soffice --headless --convert-to docx --outdir " . escapeshellarg($outDir) . " " . escapeshellarg($tmpCopy) . " 2>&1";
            exec($cmd, $out, $ret);
            $converted = $outDir . '/' . pathinfo($tmpCopy, PATHINFO_FILENAME) . '.docx';

            if ($ret === 0 && file_exists($converted) && $phpwordAvailable) {
                try {
                    $lines      = extractLinesFromDocx($converted);
                    $giftOutput = convertToGift($lines);
                } catch (Exception $e) {
                    $error = 'Error al leer el Word convertido: ' . $e->getMessage();
                }
            } else {
                $error = 'No se pudo convertir el archivo .doc/.dot. Guárdalo como .docx desde Word e inténtalo de nuevo.';
            }
        } else {
            $error = 'Solo se aceptan archivos .docx, .doc, .dot o .txt';
        }

        if ($giftOutput && substr_count($giftOutput, "\n=") === 0) {
            $warning = 'No se detectaron respuestas correctas. Asegúrate de poner una C al final de la opción correcta. Ej: "o  Texto de la respuesta C"';
        }

    } else {
        $error = 'No se pudo leer el archivo. Intenta de nuevo.';
    }
}
?>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>GIFT Converter</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Syne:wght@400;700;800&display=swap" rel="stylesheet">
<style>
  :root {
    --bg:      #0a0a0f;
    --surface: #12121a;
    --border:  #2a2a3a;
    --accent:  #7cffd4;
    --accent2: #ff6b6b;
    --warn:    #ffd97c;
    --text:    #e8e8f0;
    --muted:   #6b6b8a;
    --mono:    'Space Mono', monospace;
    --sans:    'Syne', sans-serif;
  }
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    background: var(--bg); color: var(--text);
    font-family: var(--sans); min-height: 100vh;
    display: flex; flex-direction: column; align-items: center;
    padding: 2rem 1rem 4rem; position: relative; overflow-x: hidden;
  }
  body::before {
    content: ''; position: fixed; inset: 0;
    background-image:
      linear-gradient(rgba(12, 192, 132, 0.03) 1px, transparent 1px),
      linear-gradient(90deg, rgba(50, 226, 168, 0.03) 1px, transparent 1px);
    background-size: 40px 40px; pointer-events: none; z-index: 0;
  }
  body::after {
    content: ''; position: fixed; top: -20%; left: 50%;
    transform: translateX(-50%); width: 600px; height: 400px;
    background: radial-gradient(ellipse, rgba(124,255,212,.07) 0%, transparent 70%);
    pointer-events: none; z-index: 0;
  }
  .wrapper { position: relative; z-index: 1; width: 100%; max-width: 860px; }

  header { text-align: center; margin-bottom: 3rem; }
  .badge {
    display: inline-block; font-family: var(--mono); font-size: .7rem;
    letter-spacing: .15em; color: var(--accent); border: 1px solid var(--accent);
    padding: .25rem .75rem; border-radius: 2px; margin-bottom: 1.2rem; text-transform: uppercase;
  }
  h1 { font-size: clamp(2rem,5vw,3.5rem); font-weight: 800; line-height: 1.05; letter-spacing: -.02em; }
  h1 span { color: var(--accent); }
  .subtitle { margin-top: .75rem; color: var(--muted); font-family: var(--mono); font-size: .85rem; }

  .card { background: var(--surface); border: 1px solid var(--border); border-radius: 8px; padding: 2rem; margin-bottom: 2rem; }

  .drop-zone {
    border: 2px dashed var(--border); border-radius: 6px; padding: 3rem 2rem;
    text-align: center; cursor: pointer; transition: border-color .2s, background .2s; position: relative;
  }
  .drop-zone:hover, .drop-zone.drag-over { border-color: var(--accent); background: rgba(124,255,212,.04); }
  .drop-zone input[type="file"] { position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%; }
  .drop-icon { font-size: 2.5rem; margin-bottom: .75rem; display: block; }
  .drop-label { font-family: var(--mono); font-size: .9rem; color: var(--muted); }
  .drop-label strong { color: var(--accent); font-weight: 700; }
  #file-name { margin-top: 1rem; font-family: var(--mono); font-size: .8rem; color: var(--accent); min-height: 1.2em; }

  .tip {
    margin-top: 1.25rem; background: rgba(124,255,212,.05);
    border: 1px solid rgba(124,255,212,.15); border-radius: 6px;
    padding: .85rem 1rem; font-family: var(--mono); font-size: .78rem;
    color: var(--muted); line-height: 1.9;
  }
  .tip strong { color: var(--accent); }
  .tip code {
    background: rgba(124,255,212,.1); padding: .15rem .45rem;
    border-radius: 3px; color: var(--accent); letter-spacing: .03em;
  }

  .btn {
    display: inline-flex; align-items: center; gap: .5rem;
    background: var(--accent); color: #0a0a0f;
    font-family: var(--mono); font-size: .85rem; font-weight: 700;
    letter-spacing: .05em; border: none; border-radius: 4px;
    padding: .75rem 2rem; cursor: pointer; margin-top: 1.5rem;
    width: 100%; justify-content: center; text-transform: uppercase;
    transition: transform .15s, box-shadow .15s;
  }
  .btn:hover { transform: translateY(-2px); box-shadow: 0 8px 24px rgba(124,255,212,.25); }
  .btn:active { transform: translateY(0); }

  .error-box {
    background: rgba(255,107,107,.08); border: 1px solid var(--accent2);
    color: var(--accent2); font-family: var(--mono); font-size: .85rem;
    padding: .85rem 1rem; border-radius: 4px; margin-bottom: 1.5rem;
  }
  .warn-box {
    background: rgba(255,217,124,.06); border: 1px solid var(--warn);
    color: var(--warn); font-family: var(--mono); font-size: .8rem;
    padding: .85rem 1rem; border-radius: 4px; margin-bottom: 1rem;
  }

  .output-header {
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 1rem; flex-wrap: wrap; gap: .75rem;
  }
  .output-title { font-family: var(--mono); font-size: .8rem; letter-spacing: .1em; text-transform: uppercase; color: var(--accent); }
  .btn-sm {
    background: transparent; border: 1px solid var(--accent); color: var(--accent);
    font-family: var(--mono); font-size: .75rem; padding: .4rem 1rem;
    border-radius: 4px; cursor: pointer; text-transform: uppercase;
    letter-spacing: .05em; transition: background .15s;
  }
  .btn-sm:hover { background: rgba(124,255,212,.1); }

  textarea.output {
    width: 100%; min-height: 420px; background: #0d0d14;
    border: 1px solid var(--border); border-radius: 6px; color: #b8fce8;
    font-family: var(--mono); font-size: .78rem; line-height: 1.7;
    padding: 1.25rem; resize: vertical; outline: none; caret-color: var(--accent);
  }
  textarea.output:focus { border-color: rgba(124,255,212,.4); }

  .stats { display: flex; gap: 1.5rem; flex-wrap: wrap; margin-top: 1rem; }
  .stat { font-family: var(--mono); font-size: .75rem; color: var(--muted); }
  .stat span { color: var(--accent); font-weight: 700; }

  #toast {
    position: fixed; bottom: 2rem; right: 2rem;
    background: var(--accent); color: #0a0a0f;
    font-family: var(--mono); font-size: .8rem; font-weight: 700;
    padding: .65rem 1.25rem; border-radius: 4px;
    opacity: 0; transform: translateY(10px);
    transition: opacity .3s, transform .3s;
    pointer-events: none; z-index: 100;
  }
  #toast.show { opacity: 1; transform: translateY(0); }
</style>
</head>
<body>
<div class="wrapper">

  <header>
    <div class="badge">Herramienta de conversión</div>
    <h1>GIFT<span>Converter</span></h1>
    <p class="subtitle">// Sube tu Word con la C marcada → obtén formato GIFT</p>
  </header>

  <?php if ($error): ?>
    <div class="error-box">⚠ <?= htmlspecialchars($error) ?></div>
  <?php endif; ?>

  <div class="card">
    <form method="POST" enctype="multipart/form-data" id="upload-form">
      <div class="drop-zone" id="drop-zone">
        <input type="file" name="archivo" id="archivo" accept=".docx,.doc,.dot,.txt">
        <span class="drop-icon">📄</span>
        <div class="drop-label">
          <strong>Haz clic</strong> o arrastra tu archivo aquí<br>
          Acepta <strong>.docx</strong> · <strong>.doc</strong> · <strong>.dot</strong> · <strong>.txt</strong>
        </div>
        <div id="file-name"></div>
      </div>



      <button type="submit" class="btn">⚡ Convertir a GIFT</button>
    </form>
  </div>

  <?php if ($warning): ?>
    <div class="warn-box">⚠ <?= htmlspecialchars($warning) ?></div>
  <?php endif; ?>

  <?php if ($giftOutput): ?>
  <div class="card">
    <div class="output-header">
      <div class="output-title">// Resultado GIFT</div>
      <div style="display:flex;gap:.5rem;flex-wrap:wrap;">
        <button class="btn-sm" onclick="copyOutput()">Copiar</button>
        <button class="btn-sm" onclick="downloadOutput()">Descargar .txt</button>
      </div>
    </div>

    <textarea class="output" id="gift-output" readonly><?= htmlspecialchars($giftOutput) ?></textarea>

    <?php
      $qCount  = substr_count($giftOutput, '// Start of question:');
      $correct = substr_count($giftOutput, "\n=");
      $lines   = substr_count($giftOutput, "\n");
    ?>
    <div class="stats">
      <div class="stat">Preguntas: <span><?= $qCount ?></span></div>
      <div class="stat">Respuestas correctas: <span><?= $correct ?></span></div>
      <div class="stat">Líneas: <span><?= $lines ?></span></div>
    </div>
  </div>
  <?php endif; ?>

</div>

<div id="toast">¡Copiado al portapapeles!</div>

<script>
  document.getElementById('archivo').addEventListener('change', function() {
    const name = this.files[0] ? this.files[0].name : '';
    document.getElementById('file-name').textContent = name ? '📎 ' + name : '';
  });

  const zone = document.getElementById('drop-zone');
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) {
      const dt = new DataTransfer();
      dt.items.add(file);
      document.getElementById('archivo').files = dt.files;
      document.getElementById('file-name').textContent = '📎 ' + file.name;
    }
  });

  function copyOutput() {
    const ta = document.getElementById('gift-output');
    ta.select();
    navigator.clipboard.writeText(ta.value).then(() => showToast());
  }

  function downloadOutput() {
    const text = document.getElementById('gift-output').value;
    const a = document.createElement('a');
    a.href = 'data:text/plain;charset=utf-8,' + encodeURIComponent(text);
    a.download = 'preguntas_gift.txt';
    a.click();
  }

  function showToast() {
    const t = document.getElementById('toast');
    t.classList.add('show');
    setTimeout(() => t.classList.remove('show'), 2200);
  }
</script>
</body>
</html>