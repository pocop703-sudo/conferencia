import type { ManifestStore } from '../types';

export const normalizeCode = (code: string | number): string => {
  const parsed = parseInt(String(code).trim(), 10);
  return Number.isNaN(parsed) ? String(code).trim() : String(parsed);
};

export const DEMO_MANIFEST_TEXT = `MD NFe Origem Destino Volumes Peso
3115890 3367531-1 0875 - EMPREENDIMENTOS PAGUE MENOS SA 0811 - EMPREENDIMENTOS PAGUE MENOS SA 21 86.89
3115632 3367230-1 0875 - EMPREENDIMENTOS PAGUE MENOS SA 1305 - EMPREENDIMENTOS PAGUE MENOS SA 18 52.21
3115633 3367103-1 0875 - EMPREENDIMENTOS PAGUE MENOS SA 1305 - EMPREENDIMENTOS PAGUE MENOS SA 1 2.40
3115634 3367212-1 0875 - EMPREENDIMENTOS PAGUE MENOS SA 1322 - EMPREENDIMENTOS PAGUE MENOS SA 15 57.21
3115699 3367393-1 0875 - EMPREENDIMENTOS PAGUE MENOS SA 1362 - EMPREENDIMENTOS PAGUE MENOS SA 18 84.75
3115705 3367264-1 0875 - EMPREENDIMENTOS PAGUE MENOS SA 9999 - LOJA FICTICIA SA 14 89.09`;

export function parseManifestText(rawText: string): ManifestStore[] {
  const lines = rawText
    .split(/\r?\n/)
    .map((line) => line.replace(/\s+/g, ' ').trim())
    .filter(Boolean);

  const matched: Array<{ code: string; volumes: number }> = [];
  const pendingCodes: Array<{ code: string }> = [];
  const pendingVolumes: number[] = [];
  const continuationRegex = /^(?:\d{5,9}-?\s+)?(\d{1,4})\s+\d{1,5}[.,]\d{1,3}$/;

  for (const line of lines) {
    if (/^MD\b/i.test(line)) continue;

    const storeMatches = [...line.matchAll(/(?:^|\s)(\d{3,4})\s*-\s*(?=[A-Za-zÀ-ú])/g)]
      .map((m) => m[1])
      .filter(Boolean);

    if (storeMatches.length === 0) {
      const contMatch = line.match(continuationRegex);
      if (contMatch?.[1]) {
        const vols = parseInt(contMatch[1], 10);
        const pending = pendingCodes.shift();
        if (pending) {
          matched.push({ code: pending.code, volumes: vols });
        } else {
          pendingVolumes.push(vols);
        }
      }
      continue;
    }

    const destCode = storeMatches[storeMatches.length - 1];
    const afterCode = line.slice(line.lastIndexOf(destCode) + destCode.length);
    const numsAfter = afterCode.match(/\d+(?:[.,]\d+)?/g) ?? [];

    if (storeMatches.length === 1 && numsAfter.length >= 3) {
      continue;
    }

    let parsedVol: number | null = null;
    const endMatch = line.match(/(\d{1,4})\s+\d{1,5}(?:[.,]\d+)\s*$/);
    if (endMatch?.[1]) {
      parsedVol = parseInt(endMatch[1], 10);
    } else {
      const words = [...afterCode.matchAll(/\b(\d{1,4})\b/g)].map((m) => parseInt(m[1] ?? '', 10));
      if (words.length && Number.isFinite(words[0])) {
        parsedVol = words[0];
      }
    }

    if (parsedVol === null) {
      const orphanVol = pendingVolumes.shift();
      if (orphanVol === undefined) {
        pendingCodes.push({ code: destCode });
      } else {
        matched.push({ code: destCode, volumes: orphanVol });
      }
      continue;
    }

    if (Number.isFinite(parsedVol) && parsedVol >= 0) {
      matched.push({ code: destCode, volumes: parsedVol });
    }
  }

  // Aggregate by normalized store code
  const storeMap = new Map<string, ManifestStore>();
  for (const item of matched) {
    const key = normalizeCode(item.code);
    const existing = storeMap.get(key);

    if (item.volumes <= 1) {
      if (existing) {
        existing.ignored += 1;
      } else {
        storeMap.set(key, { code: item.code, volumes: 0, occurrences: 0, ignored: 1 });
      }
      continue;
    }

    if (!existing || existing.volumes === 0) {
      storeMap.set(key, {
        code: item.code,
        volumes: item.volumes,
        occurrences: 1,
        ignored: existing?.ignored ?? 0,
      });
    } else {
      existing.occurrences += 1;
      existing.volumes = Math.max(existing.volumes, item.volumes);
    }
  }

  return [...storeMap.values()]
    .filter((e) => e.volumes > 1)
    .sort((a, b) => normalizeCode(a.code).localeCompare(normalizeCode(b.code), 'pt-BR', { numeric: true }));
}

async function extractTextFromPdf(
  file: File,
  onProgress: (msg: string) => void
): Promise<{ text: string; pages: number; empty: boolean }> {
  const pdfjsLib = await import('pdfjs-dist');
  try {
    if (!pdfjsLib.GlobalWorkerOptions.workerSrc) {
      pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;
    }
  } catch {
    // fallback
  }

  const arrayBuffer = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({
    data: new Uint8Array(arrayBuffer),
    cMapUrl: `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/cmaps/`,
    cMapPacked: true,
  });

  const pdfDoc = await loadingTask.promise;
  const lines: string[] = [];

  for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
    onProgress(`Lendo página ${pageNum} de ${pdfDoc.numPages}...`);
    const page = await pdfDoc.getPage(pageNum);
    const content = await page.getTextContent();

    const rowMap = new Map<number, Array<{ x: number; s: string }>>();
    for (const item of content.items) {
      if ('str' in item && typeof item.str === 'string' && item.str.trim()) {
        const y = Math.round((item.transform[5] as number) / 4) * 4;
        const row = rowMap.get(y) ?? [];
        row.push({ x: item.transform[4] as number, s: item.str });
        rowMap.set(y, row);
      }
    }

    const sortedY = [...rowMap.keys()].sort((a, b) => b - a);
    for (const y of sortedY) {
      const line = rowMap
        .get(y)!
        .sort((a, b) => a.x - b.x)
        .map((i) => i.s)
        .join(' ');
      lines.push(line);
    }
  }

  const fullText = lines.join('\n');
  const isEmpty = fullText.replace(/\s/g, '').length < 40;
  return {
    text: fullText,
    pages: pdfDoc.numPages,
    empty: isEmpty,
  };
}

async function runOcr(file: File | Blob, onProgress: (msg: string) => void): Promise<string> {
  onProgress('Executando OCR no documento (Tesseract)...');
  const { createWorker } = await import('tesseract.js');
  const worker = await createWorker('por');
  const ret = await worker.recognize(file);
  await worker.terminate();
  return ret.data.text;
}

export async function readManifestFile(
  file: File,
  onProgress: (msg: string) => void = () => {}
): Promise<string> {
  if (file.type === 'application/pdf' || /\.pdf$/i.test(file.name)) {
    try {
      const { text, empty } = await extractTextFromPdf(file, onProgress);
      if (empty) {
        onProgress('PDF sem camada de texto — iniciando OCR...');
        return await runOcr(file, onProgress);
      }
      return text;
    } catch (err: unknown) {
      console.warn('PDF reader error, fallback to OCR:', err);
      onProgress('Tentando leitura via OCR...');
      return await runOcr(file, onProgress);
    }
  }
  return await runOcr(file, onProgress);
}
