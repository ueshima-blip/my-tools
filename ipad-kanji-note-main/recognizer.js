/* recognizer.js
 * 漢字（ひらがな・カタカナを含む）の手書き認識を、システム搭載の
 * 日本語フォントで描画したリファレンス画像とのテンプレートマッチで実装する。
 *
 * 入力: ストローク列（[{points:[{x,y,p,t}, ...]}, ...]）
 * 出力: { best, bestScore, targetScore, top: [[char, score], ...] }
 *
 * 仕組み:
 *  1. 描かれたストロークを 64x64 の画像に正規化（重心ではなく外接矩形を 80% 以下に収める）
 *  2. Sobel 風の軽いガウシアンぼかしで多少のずれを許容
 *  3. リファレンス候補（お題＋まぎらわしい字＋小学校で習う頻出字）をフォントで描画してキャッシュ
 *  4. コサイン類似度（中心化なし）で比較してスコア化
 *
 * 完璧な OCR ではないが、「お題と違う字を書いている」かどうかを判定するには十分。
 * 候補の中で最大スコアの字 ≠ お題、かつ差が一定以上のとき「ちがう、正しくは X」と返す。
 */
(function (root) {
  'use strict';

  // --------- 候補リスト ----------

  // よく混同するペア／グループ。お題の字がここに含まれるとき、同グループの全字を強制的に候補に入れる。
  const CONFUSION_GROUPS = [
    ['日','月','目','白','百','自'],
    ['人','入','八','大','犬','太','天','夫'],
    ['本','木','末','未','体'],
    ['土','士','王','玉','主'],
    ['右','左','石','名','各'],
    ['田','由','申','甲','男'],
    ['上','下','卜'],
    ['子','字','学'],
    ['校','枝','枚','林','森'],
    ['花','化','茶'],
    ['犬','大','太'],
    ['口','回','日','目'],
    ['力','刀','九'],
    ['先','光','見'],
    ['年','午','牛','生'],
    ['川','三','小'],
    ['東','車','束'],
    ['南','西','北'],
    ['雨','両','面'],
    ['手','毛'],
    ['白','百'],
    ['千','干','于'],
    ['名','多','夕'],
    ['早','旱','旦'],
    ['正','止'],
    ['赤','亦'],
    ['青','晴','清'],
    ['空','穴','究'],
    ['気','汽'],
    ['音','立','音'],
    ['出','山'],
    ['村','付','寸'],
  ];

  // 小学校１・２年生で習う漢字（おおむね）＋かな。候補プールに常に含める。
  const COMMON_CHARS = (
    // 1年
    '一二三四五六七八九十百千月日年上下左右大中小本人入木林森山川水火土金石田力口耳目手足生先名学校字玉王糸貝音車赤青白花草虫犬猫魚鳥太天空雨気早夕町村休見立文書読新春夏秋冬東西南北円黒' +
    // 2年（抜粋）
    '父母兄弟姉妹友前後外内自分体頭顔首肩声話言語食事毎朝晩昼夜時間今週曜風雪雲海岸寺池谷岩谷麦米茶肉魚鳥馬牛羊豚汽船工場黄黒紙地図形角線点番算数理科社会国家民間家族' +
    // よく書く字
    'はがき手紙電話学校先生友達犬猫鉛筆消'
  );

  function uniqueChars(s) {
    const seen = new Set();
    const out = [];
    for (const c of s) if (!seen.has(c) && c.trim()) { seen.add(c); out.push(c); }
    return out;
  }
  const CHAR_POOL = uniqueChars(COMMON_CHARS);

  // --------- 画像化ユーティリティ ----------

  const SIZE = 64;       // 認識用ビットマップサイズ
  const MARGIN = 0.10;   // 内側余白（割合）

  /** 候補漢字を SIZE×SIZE のグレースケール（インク濃度 0..1）配列に */
  function renderRefBitmap(char) {
    const c = document.createElement('canvas');
    c.width = SIZE; c.height = SIZE;
    const x = c.getContext('2d', { willReadFrequently: true });
    x.fillStyle = '#fff';
    x.fillRect(0, 0, SIZE, SIZE);
    x.fillStyle = '#000';
    x.textAlign = 'center';
    x.textBaseline = 'middle';
    // 楷書系の見え方に近い明朝/ゴシックを使い、外接が大きい方の字でも見切れないよう少し小さめに
    const fontSize = Math.floor(SIZE * 0.78);
    x.font = `${fontSize}px "Hiragino Mincho ProN","YuMincho","Yu Mincho","Noto Serif JP","Hiragino Sans","Yu Gothic","Noto Sans JP",serif`;
    // 漢字はアセンダ/ディセンダがほぼ同じだが、安全のため微調整
    x.fillText(char, SIZE / 2, SIZE / 2 + Math.floor(SIZE * 0.04));
    const img = x.getImageData(0, 0, SIZE, SIZE);
    return imageDataToInk(img);
  }

  function imageDataToInk(imgData) {
    const { data, width, height } = imgData;
    const out = new Float32Array(width * height);
    for (let i = 0; i < width * height; i++) {
      const r = data[i * 4], g = data[i * 4 + 1], b = data[i * 4 + 2], a = data[i * 4 + 3];
      // 黒インク濃度: 不透明度 × (1 - 明度)
      const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
      out[i] = (a / 255) * (1 - lum);
    }
    return out;
  }

  /** 3x3 ボックスフィルタ（軽量ぼかし）。位置のずれを吸収して類似度を安定化させる。 */
  function blur3(src, w, h) {
    const out = new Float32Array(src.length);
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        let s = 0, n = 0;
        for (let dy = -1; dy <= 1; dy++) {
          const yy = y + dy;
          if (yy < 0 || yy >= h) continue;
          for (let dx = -1; dx <= 1; dx++) {
            const xx = x + dx;
            if (xx < 0 || xx >= w) continue;
            s += src[yy * w + xx];
            n++;
          }
        }
        out[y * w + x] = s / n;
      }
    }
    return out;
  }

  /** ストローク群を SIZE×SIZE のインク濃度配列に。`bbox` 指定があればその矩形を入力空間として使う。 */
  function strokesToBitmap(strokes, opts = {}) {
    if (!strokes || strokes.length === 0) return null;

    let minX, minY, maxX, maxY;
    if (opts.bbox) {
      ({ minX, minY, maxX, maxY } = opts.bbox);
    } else {
      minX = Infinity; minY = Infinity; maxX = -Infinity; maxY = -Infinity;
      for (const s of strokes) for (const p of s.points) {
        if (p.x < minX) minX = p.x;
        if (p.y < minY) minY = p.y;
        if (p.x > maxX) maxX = p.x;
        if (p.y > maxY) maxY = p.y;
      }
    }
    const bw = maxX - minX, bh = maxY - minY;
    if (!isFinite(bw) || !isFinite(bh)) return null;
    if (bw <= 0 && bh <= 0) return null;

    // bbox の中心を SIZE/2 に置き、長辺を `inner` に合わせて拡縮（縦横比は維持）
    const inner = SIZE * (1 - MARGIN * 2);
    const longSide = Math.max(bw, bh, 1);
    const scale = inner / longSide;
    const cx = (minX + maxX) / 2, cy = (minY + maxY) / 2;

    const c = document.createElement('canvas');
    c.width = SIZE; c.height = SIZE;
    const x = c.getContext('2d', { willReadFrequently: true });
    x.fillStyle = '#fff';
    x.fillRect(0, 0, SIZE, SIZE);
    x.strokeStyle = '#000';
    x.lineCap = 'round';
    x.lineJoin = 'round';
    // 線の太さは入力 bbox の規模に追従（拡大した結果として一定の太さに見えるように）
    const drawnThickness = (opts.refThickness != null)
      ? opts.refThickness
      : Math.max(2, longSide * 0.06) * scale;
    x.lineWidth = drawnThickness;

    for (const s of strokes) {
      const pts = s.points;
      if (!pts || pts.length === 0) continue;
      x.beginPath();
      for (let i = 0; i < pts.length; i++) {
        const px = (pts[i].x - cx) * scale + SIZE / 2;
        const py = (pts[i].y - cy) * scale + SIZE / 2;
        if (i === 0) x.moveTo(px, py); else x.lineTo(px, py);
      }
      // 1点だけのストロークは小さい点として描画
      if (pts.length === 1) {
        const px = (pts[0].x - cx) * scale + SIZE / 2;
        const py = (pts[0].y - cy) * scale + SIZE / 2;
        x.moveTo(px + 0.5, py);
        x.lineTo(px - 0.5, py);
      }
      x.stroke();
    }
    const img = x.getImageData(0, 0, SIZE, SIZE);
    return imageDataToInk(img);
  }

  // --------- 類似度 ----------

  /** L2 正規化したコサイン類似度（>=0） */
  function cosine(a, b) {
    let dot = 0, na = 0, nb = 0;
    const n = a.length;
    for (let i = 0; i < n; i++) {
      const va = a[i], vb = b[i];
      dot += va * vb;
      na += va * va;
      nb += vb * vb;
    }
    if (na === 0 || nb === 0) return 0;
    return dot / Math.sqrt(na * nb);
  }

  // --------- リファレンスキャッシュ ----------

  const refCache = new Map();
  function getRef(char) {
    if (refCache.has(char)) return refCache.get(char);
    const raw = renderRefBitmap(char);
    const blurred = blur3(raw, SIZE, SIZE);
    refCache.set(char, blurred);
    return blurred;
  }

  // --------- 公開 API ----------

  /**
   * @param {Stroke[]} strokes
   * @param {Object} [options]
   * @param {string} [options.target]      期待する字。あれば候補に強制追加し targetScore を返す。
   * @param {string[]} [options.candidates] 追加の候補。
   * @param {{minX,minY,maxX,maxY}} [options.bbox] 入力空間の矩形。指定が無ければストロークの外接矩形。
   * @returns {{best:string, bestScore:number, targetScore:number, top:[string,number][]} | null}
   */
  function recognize(strokes, options = {}) {
    const drawn = strokesToBitmap(strokes, { bbox: options.bbox });
    if (!drawn) return null;
    const drawnB = blur3(drawn, SIZE, SIZE);

    const candidates = new Set();
    if (options.target) {
      for (const c of options.target) candidates.add(c); // 目標字
      for (const ch of options.target) {
        for (const g of CONFUSION_GROUPS) if (g.includes(ch)) for (const c of g) candidates.add(c);
      }
    }
    if (options.candidates) for (const c of options.candidates) candidates.add(c);
    for (const c of CHAR_POOL) candidates.add(c);

    const scores = [];
    let bestChar = null, bestScore = -1, targetScore = -1;
    for (const ch of candidates) {
      const ref = getRef(ch);
      const s = cosine(drawnB, ref);
      scores.push([ch, s]);
      if (s > bestScore) { bestScore = s; bestChar = ch; }
      if (options.target && ch === options.target) targetScore = s;
    }
    scores.sort((a, b) => b[1] - a[1]);
    return {
      best: bestChar,
      bestScore,
      targetScore,
      top: scores.slice(0, 8),
    };
  }

  /**
   * 判定の意味づけ。返り値: 'ok' | 'ng' | 'unsure'
   *   ok      : 正しい字を書けている
   *   ng      : ちがう字（best が target と異なる）
   *   unsure  : スコアが低すぎてはっきりしない（書きかけ等）
   */
  function classify(result, target) {
    if (!result) return 'unsure';
    const { best, bestScore, targetScore } = result;
    // 全体的にスコアが低すぎる: 書きかけ／ぐちゃぐちゃ
    if (bestScore < 0.42) return 'unsure';
    // best が target と一致 → ok
    if (best === target && targetScore >= 0.42) return 'ok';
    // best が target と異なるが、target スコアが best とほぼ同じ → ok 寄りで unsure
    if (best !== target && targetScore >= 0.42 && (bestScore - targetScore) < 0.03) return 'ok';
    // 明確に best が別の字 → ng
    if (best !== target && bestScore - targetScore >= 0.03) return 'ng';
    return 'unsure';
  }

  root.KanjiRecognizer = {
    recognize,
    classify,
    strokesToBitmap,
    getRef,
    SIZE,
  };
})(window);
