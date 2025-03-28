/**
 * Markdown形式のテキストからスライドを生成します。
 *
 * parseMarkdownToSlides() で解析した結果を元に、以下の手順でスライドを作成します：
 * 1. プレゼンテーション全体のタイトルを設定するタイトルスライドを追加
 * 2. 各スライドデータ（タイトル、本文（箇条書き）、コードブロックなど）に対応して、スライドを追加
 * 3. 生成したスライドを、現在選択中のスライドの後ろに移動
 *
 * @param {string} markdown - Markdown形式の入力テキスト
 */
const createSlidesFromMarkdown = (markdown) => {
    // Markdownを解析して、タイトルと各スライドのデータを取得
    const parsedData = parseMarkdownToSlides(markdown);
    const presentation = SlidesApp.getActivePresentation();

    // 現在選択中のスライドの位置を取得（既存の関数を利用）
    const selection = presentation.getSelection();
    let slideIndex = 0;
    try {
        slideIndex = getSlideIndexByObjectId_(selection.getCurrentPage().getObjectId());
    } catch (e) {
        // 選択情報が取得できない場合は末尾に追加
        slideIndex = presentation.getSlides().length - 1;
    }

    const addedSlides = [];

    // プレゼンテーションタイトル用のスライドを作成
    const titleSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE);
    addedSlides.push(titleSlide);
    titleSlide.getShapes()[0].getText().setText(parsedData.title);

    // 各スライドのデータに基づいて、スライドを生成
    parsedData.slides.forEach(slideData => {
        // 使用するレイアウトを判定
        const layout = slideData.sectionHeader 
            ? SlidesApp.PredefinedLayout.SECTION_HEADER 
            : SlidesApp.PredefinedLayout.TITLE_AND_BODY;
        const slide = presentation.appendSlide(layout);
        addedSlides.push(slide);

        // スライドタイトルを設定
        slide.getShapes()[0].getText().setText(slideData.title);

        // 本文（箇条書き）の設定：
        // 各箇条書き行を段落として追加し、箇条書きスタイルを適用
        const bodyShape = slide.getShapes()[1];
        slideData.points.forEach(point => {
            bodyShape.getText().appendParagraph(point)
                .getRange()
                .getListStyle()
                .applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);
        });

        // ※コードブロックがある場合は、既存の insertCodeBlockText_() などで処理可能
        if (slideData.code) {
            insertCodeBlockText_(slideData.code, slide);
        }
    });

    // 生成したスライドを、選択中のスライドの後ろに移動する
    for (let i = addedSlides.length - 1; i >= 0; i--) {
        addedSlides[i].move(slideIndex + 1);
    }
};

/**
 * Markdown形式のテキストからスライドデータに変換します。
 *
 * ルール:
 * 1. 1階層（# ）の見出しはプレゼンテーション全体のタイトル（最初に出現したもの）とする。
 * 2. 2階層（## ）および3階層（### ）の見出しは「見出し」として、タイトルのみのスライドを出力する。
 * 3. 箇条書き行（"-" または "*" で始まる行）が連続する場合、
 *    ・グループ内に入れ子（indent > 0）がある場合は、インデント0の行を見出しとして新規スライドとし、
 *      入れ子をそのスライドの本文 (points) に追加する。
 *    ・グループ内がすべてインデント0の場合は、直前の見出しスライド（または、なければ新規に "内容" スライド）にまとめる。
 * 4. コードブロックは ``` で囲まれた部分をひとまとめの code として扱う。
 *
 * @param {string} markdown - Markdown形式のテキスト
 * @returns {Object} { title: string, slides: Array }
 *   ・title: プレゼンテーションのタイトル（最初の "# "）
 *   ・slides: 各スライドのデータ。各スライドは { title: string, points: Array, code?: string } の形式
 */
const parseMarkdownToSlides = (markdown) => {
  let title = "";
  const slides = [];
  let currentSlide = null;  // 通常テキストで生成するスライド用
  let inCodeBlock = false;
  let codeBuffer = "";
  const lines = markdown.split("\n");

  // 箇条書きグループと通常テキストを一時保持する配列
  let contentGroup = [];
  
  /**
   * 行頭の空白数を返します。
   *
   * @param {string} line - 対象の行
   * @returns {number} 行頭の空白の数
   */
  const getIndent = (line) => {
    const match = line.match(/^(\s*)/);
    return match ? match[1].length : 0;
  };

  /**
   * contentGroupを処理し、スライドまたは現在のスライドに項目を追加します。
   * 順序を保持しながら、階層構造がある場合は適切に処理します。
   */
  const flushContentGroup = () => {
    if (contentGroup.length === 0) return;
    
    // contentGroup 内にインデント > 0 の要素が存在するか判定（プレーンテキスト以外）
    const hierarchical = contentGroup.some(item => !item.isPlainText && item.indent > 0);

    if (hierarchical) {
      // 階層構造がある場合、インデント0の箇条書きを見出しとして扱う
      let currentBulletSlide = null;
      let foundTitle = false;
      
      contentGroup.forEach(item => {
        if (!item.isPlainText && item.indent === 0) {
          // インデント0の箇条書きは新しいスライドのタイトルになる
          if (currentBulletSlide) {
            slides.push(currentBulletSlide);
          }
          currentBulletSlide = { title: item.content, points: [] };
          foundTitle = true;
        } else {
          // それ以外はすべて本文に追加
          if (!currentBulletSlide) {
            currentBulletSlide = { title: "内容", points: [] };
          }
          currentBulletSlide.points.push(item.content);
        }
      });
      
      if (currentBulletSlide) {
        slides.push(currentBulletSlide);
      }
    } else {
      // 階層構造がない場合：直前のスライド（または現在のスライド）に追加
      let targetSlide = slides.length > 0 ? slides[slides.length - 1] : currentSlide;
      if (!targetSlide) {
        targetSlide = { title: "内容", points: [] };
        currentSlide = targetSlide;
        slides.push(currentSlide);
      }
      
      contentGroup.forEach(item => {
        targetSlide.points.push(item.content);
      });
    }
    
    contentGroup = [];
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmed = line.trim();

    // 空行の場合はスキップ
    if (trimmed === "") {
      continue;
    }

    // コードブロックの開始／終了判定
    if (trimmed.startsWith("```")) {
      flushContentGroup();
      inCodeBlock = !inCodeBlock;
      if (!inCodeBlock && currentSlide && codeBuffer) {
        currentSlide.code = codeBuffer.trim();
        codeBuffer = "";
      }
      continue;
    }
    
    if (inCodeBlock) {
      codeBuffer += line + "\n";
      continue;
    }

    // 1階層の見出し：プレゼンテーション全体のタイトル
    if (trimmed.startsWith("# ") && !trimmed.startsWith("## ") && !trimmed.startsWith("### ")) {
      flushContentGroup();
      const headerText = trimmed.replace("# ", "").trim();
      if (!title) {
        title = headerText;
      } else {
        if (currentSlide) {
          slides.push(currentSlide);
          currentSlide = null;
        }
        currentSlide = { title: headerText, points: [] };
      }
      continue;
    }

    // 2階層または3階層の見出し → flush 後、タイトルのみのスライドを出力
    if (trimmed.startsWith("## ") || trimmed.startsWith("### ")) {
      flushContentGroup();
      const headerText = trimmed.replace(/^#{2,3}\s*/, "").trim();
      if (currentSlide) {
        slides.push(currentSlide);
        currentSlide = null;
      }
      slides.push({ title: headerText, points: [] });
      continue;
    }

    // スライド区切り（---）の処理
    if (trimmed === "---") {
      flushContentGroup();
      if (currentSlide) {
        currentSlide = null;
      }
      slides.push({ title: '', points: [] });
      
      continue;
    }

    // 箇条書き行の処理（"-" または "*" で始まる行）
    const bulletMatch = line.match(/^(\s*)[-*]\s+(.*)/);
    if (bulletMatch) {
      const indent = bulletMatch[1].length;
      const content = bulletMatch[2].trim();
      contentGroup.push({ indent: indent, content: content, isPlainText: false });
    } else {
      // 箇条書き以外の非空行は通常テキストとして追加
      contentGroup.push({ indent: 0, content: trimmed, isPlainText: true });
    }
  }
  
  flushContentGroup();
  if (currentSlide && !slides.includes(currentSlide)) {
    slides.push(currentSlide);
  }

  // points が空の場合、sectionHeader:true を追加
  slides.forEach(slide => {
    if (!slide.points || slide.points.length === 0) {
      slide.sectionHeader = true;
    }
  });

  return { title: title || "タイトル", slides: slides };
};
