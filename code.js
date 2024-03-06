const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
const OPENAI_MODEL = PropertiesService.getScriptProperties().getProperty("OPENAI_MODEL") || "gpt-3.5-turbo-16k";
const IMAGE_TEMPERATURE = PropertiesService.getScriptProperties().getProperty("IMAGE_TEMPERATURE") || "0.5";
const IMAGE_MAX_WORDS = PropertiesService.getScriptProperties().getProperty("IMAGE_MAX_WORDS") || "40";
const SHEET_ID = PropertiesService.getScriptProperties().getProperty("SHEET_ID");

function onOpen() {
  const menu = SlidesApp.getUi().createAddonMenu();
  menu.addItem('GPTスライド', 'showSidebar');
  menu.addToUi();
}

function showSidebar() {
  const ver = "(ver:1)";
  const html = HtmlService.createHtmlOutputFromFile('sidebar.html')
      .setTitle('GPTスライド'  + ver)
      .setWidth(300);
  SlidesApp.getUi().showSidebar(html);
}

/**
 * 現在のプレゼンテーションにスライドを作成します。
 * 
 * @param {string} inputText - スライド作成のための指示
 * @param {number} numPages - 作成するスライドの数
 * @param {string} creativity - スライド作成の創造性レベル
 * 
 * この関数は、GPTから取得したレスポンスを使用してスライドを作成します。
 * まず、GPTレスポンスからスライドデータを取得し、アクティブなプレゼンテーションを取得します。
 * 次に、タイトルスライドを作成し、タイトルテキストを設定します。
 * 最後に、各スライドデータオブジェクトをループしてスライドを作成します。
 * スライドデータに"code"フィールドがある場合、コードブロックを追加します。
 * スライドデータに"points"フィールドがある場合、箇条書きを追加します。
 * すべてのスライドが作成された後、それらを適切な位置に移動します。
 */
function createSlidesInCurrentPresentation(inputText, numPages, creativity) {
  // GPTレスポンスからJSONデータを取得
  const json = getSlidesResponse_(inputText, numPages, creativity);

  // アクティブなプレゼンテーションとスライドデータを取得
  const currentPresentation = SlidesApp.getActivePresentation();
  const slides = json.slides;

  const slideObjectId = currentPresentation.getSelection().getCurrentPage().getObjectId();
  const slideIndex = getSlideIndexByObjectId_(slideObjectId);
  const addedSlides = [];

  // タイトルスライドを作成し、タイトルテキストを設定
  const titleSlide = currentPresentation.appendSlide(SlidesApp.PredefinedLayout.TITLE);
  addedSlides.push(titleSlide);

  titleSlide.getShapes()[0].getText().setText(json.title);

  // 各スライドデータオブジェクトをループし、スライドを作成
  slides.forEach((slideData, index) => {
    // タイトルと本文のレイアウトで新しいスライドを追加
    const slide = currentPresentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
    addedSlides.push(slide);

    // スライドからタイトルと本文のShapeを取得
    const titleShape = slide.getShapes()[0];
    const bodyShape = slide.getShapes()[1];

    // スライドのタイトルを設定
    titleShape.getText().setText(slideData.title);

    // slideDataに"code"フィールドがある場合、コードブロックを追加
    if (slideData.code) {
      insertCodeBlockText_(slideData.code,slide);
    }

    // slideDataに"points"フィールドがある場合、箇条書きを追加
    if (slideData.points) {
      for (const point of slideData.points) {
        if (point.code) { // 念のため。
          insertCodeBlockText_(point.code,slide);        
        } else {
          bodyShape.getText().appendParagraph(point).getRange().getListStyle().applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);
        }
      }
    }
  });
  for (let i = addedSlides.length - 1;i >= 0;i--) {
    const slide = addedSlides[i];
    slide.move(slideIndex + 1);
  }
}

/**
 * 現在のスライドに画像を生成します。
 * 
 * @param {string} imageCaption - 生成する画像のキャプション(スタイル)
 * 
 * この関数は、アクティブなスライドから全てのテキストを取得し、画像生成のためのプロンプトを作成します。
 * 次に、Dall-Eから画像を生成し、その画像をアクティブなスライドに挿入します。
 */
function generateImageInCurrentSlide(imageCaption) {
  const allText = getAllTextFromActiveSlide_();
  const prompt = getImagePrompt_(allText);

  const blob = generateImageFromDallE_(prompt, imageCaption);
  insertImageBlobToActiveSlide_(blob)

}

/**
 * 現在選択されている要素にテキストを生成します。
 * 
 * @param {string} deepDive - 指示（深掘りするためのプロンプト）
 * @param {string} creativity - テキスト生成の創造性レベル
 * 
 * この関数は、アクティブなプレゼンテーションから選択された要素を取得し、それがテキストを含むShapeまたはテーブルである場合、
 * そのテキストを取得し、新しいテキストを生成して置き換えます。
 * テキストShapeの場合、そのテキストを取得し、新しいテキストを生成して置き換えます。
 * テーブルの場合、各セルのテキストを取得し、新しいテキストを生成して置き換えます。
 * 選択された要素がテキストを含む形状またはテーブルでない場合、エラーをスローします。
 */
function generateTextInCurrentElement(deepDive, creativity) {
  const selection = SlidesApp.getActivePresentation().getSelection();
  if (!selection || !selection.getPageElementRange()) {
    throw "テキストが含まれるオブジェクトをアクティブにしてください。";
  }
  const selectedElements = selection.getPageElementRange().getPageElements();
  
  selectedElements.forEach(element => {
    if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE && element.asShape().getText()) {
      const shape = element.asShape();
      const textRange = shape.getText();
      const currentText = textRange.asString();
      const content = getTextResponse_(deepDive + "\n" +currentText, creativity);
      const newText = `${content}`;
      textRange.setText(newText);
    } else if (element.getPageElementType() === SlidesApp.PageElementType.TABLE) {
      const table = element.asTable();
      const numRows = table.getNumRows();
      const numCols = table.getNumColumns();
      // 最初のカラムはヘッダであることを前提とする
      for (let i = 1; i < numRows; i++) {
        for (let j = 0; j < numCols; j++) {
          const cell = table.getCell(i, j);
          const textRange = cell.getText();
          const currentText = textRange.asString();
          const content = getTextResponse_(deepDive + "\n" +currentText, creativity);
          const newText = `${content}`;
          textRange.setText(newText);
        }
      }
    }
  });
}

function getSlideIndexByObjectId_(objectId)  {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  
  const index = slides.findIndex(slide => slide.getObjectId() === objectId);
  return index === -1 ? 0 : index;
}

function getTextResponse_(prompt, creativity) {
  let temperature = 0.5;
  if (creativity === "low") {
    temperature = 0.0;
  } else if (creativity === "high") {
    temperature = 1.0;
  }

  const url = "https://api.openai.com/v1/chat/completions";
  const data = {
    model: OPENAI_MODEL,
    messages: [
      { role: "user", content: prompt }
    ],
    max_tokens: 1000,
    temperature: temperature
  };

  const options = {
    method: "post",
    headers: {
      Authorization: "Bearer " + OPENAI_API_KEY,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  };

  log_("【REQ_3】" + JSON.stringify(data));

  const response = UrlFetchApp.fetch(url, options);
  const jsonResponse = JSON.parse(response.getContentText());
  const content = jsonResponse.choices[0].message.content.trim();
  log_("【RES_3】" + content);
  return content;
}

function getAllTextFromActiveSlide_() {
  // Google スライドのアクティブなプレゼンテーションとスライドを取得
  const presentation = SlidesApp.getActivePresentation();
  const slide = presentation.getSelection().getCurrentPage();

  // スライド内のすべてのページ要素を取得
  const pageElements = slide.getPageElements();
  let allText = '';

  // ページ要素をループしてテキストボックスを見つける
  for (const pageElement of pageElements) {
    // テキストボックスの場合、テキストを取得して結果に追加
    if (pageElement.getPageElementType() === SlidesApp.PageElementType.SHAPE && pageElement.asShape().getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
      const text = pageElement.asShape().getText().asString();
      allText += text + '\n';
    }
  }

  return allText;

}/**
 * OpenAI APIを使用してスライドのレスポンスを取得します。
 * 
 * @param {string} prompt - スライド生成のためのプロンプト
 * @param {number} numPages - 生成するスライドの数
 * @param {string} creativity - スライド生成の創造性レベル
 * @returns {Object} スライドデータを含むJSONオブジェクト
 * 
 * この関数は、OpenAI APIを使用してスライドのレスポンスを取得します。
 * システムメッセージは、スライドのアウトラインを作成する指示とテンプレートJSONを含みます。
 * ユーザーメッセージは、プロンプトを含みます。
 * レスポンスフォーマットは、タイプを"json_object"に指定していますが、
 * 念のため、JSONコードブロックを削除する従来の処理も行っています。
 * パースに失敗した場合、エラーをログに記録し、エラーをスローします。
 */
function getSlidesResponse_(prompt, numPages, creativity) {
  let temperature = 0.5;
  if (creativity === "low") {
    temperature = 0.0;
  } else if (creativity === "high") {
    temperature = 1.0;
  }
  const url = "https://api.openai.com/v1/chat/completions";
  const templateJson = {
    "title": "タイトル",
    "slides":[
      {"title": "スライドタイトル",
      "code": "programing code",
      "points": [" 項目 ", " 項目 "," 項目 "]
      }
    ] 
  };
  const data = {
    model: OPENAI_MODEL,
    messages: [
      {
        role: "system",
        content: `Create a detailed outline with titles and bullet points required to explain the topic. you must output the result in the following JSON format. make sure the number of elements in the "slides" array is ${numPages}. if you generate source code ,put in the "code" element. you must enclose each JSON element in double quotes ("").\n ${JSON.stringify(templateJson)}\n `
      },
      { role: "user", content: prompt }
    ],
    // max_tokens: 3000,
    temperature: temperature,
    response_format: {
      type: "json_object"
    }
  };

  const options = {
    method: "post",
    headers: {
      Authorization: "Bearer " + OPENAI_API_KEY,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  };

  log_("【REQ_1】" + JSON.stringify(data));

  const response = UrlFetchApp.fetch(url, options);
  const jsonResponse = JSON.parse(response.getContentText());
  const content = jsonResponse.choices[0].message.content.trim();
  log_("【RES_1】" + content);

  try {
    return JSON.parse(removeJsonCodeBlock_(content));
  } catch(e) {
    log_("【ERR_1】" + e);
    throw e;
  }
}
// たまに意図しない出力をすることへの対応
function removeJsonCodeBlock_(inputString) {
  const outputString = inputString.replace(/```json/g, '').replace(/```/g, '');
  return outputString;
}

/**
 * スライドにコードブロックのテキストを挿入します。
 * 
 * @param {string} codeblock - 挿入するコードブロック
 * @param {Object} slide - コードブロックを挿入するスライド
 * 
 * スライドにコードブロックのテキストを新しいテキストShapeとして挿入します。
 * コードとしての見栄えを良くするために、フォントサイズやファミリなどを独自に設定しています。
 * 
 */
function insertCodeBlockText_(codeblock,slide) {
  // 連続する複数の改行を1つの改行に置き換え
  const code = codeblock.replace(/(\n)+/g, '\n');

  // コードを行に分割し、行数に応じてフォントサイズを設定
  const texts = code.split("\n");
  const textsLen = texts.length;
  // たまに意図せず変なコードを作ることへの対応
  if (textsLen <= 1) {
    return;
  }
  const fontSize = textsLen > 20 ? 8 : 12;
  const bodyShape = slide.getShapes()[1];
  const shape = bodyShape.duplicate().setLeft(200).setWidth(500).asShape();

  // パラグラフとしてコードの行を追加し、等幅フォントと調整された間隔で表示
  const textBox = shape.getText();
  for (const text of texts) {
    const newParagraph = textBox.appendParagraph(text);
    newParagraph.getRange().getParagraphStyle().setSpaceBelow(0).setLineSpacing(110);
    newParagraph.getRange().getTextStyle().setFontSize(fontSize).setFontFamily("Courier");
  }
}

/**
 * OpenAI APIを使用して画像生成のためのプロンプトを取得します。
 * 
 * @param {string} targetText - 画像生成のためのターゲットテキスト
 * @returns {string} 画像生成AIのためのプロンプト
 * 
 * この関数は、OpenAI APIを使用して画像生成（DALL-E）のためのプロンプトを取得します。
 * 
 */
function getImagePrompt_(targetText) {
  const url = "https://api.openai.com/v1/chat/completions";

  const getPositivePrompt = function() {
    const prompt = `Please imagine a common scene from the given text and express it as a detailed prompt in english for image generation AI.\n
  rule:\n- begin  a sentence with  "The prompt is:".\n- without using bullet points.\n- in the third person.\n- ${IMAGE_MAX_WORDS} words.\ntext :`;
    return prompt;
  }

  const data = {
    model: OPENAI_MODEL,
    messages: [
      {
        role: "system",
        content: getPositivePrompt()
      },
      { role: "user", content: targetText || 'The message "error" in white wall.' }
    ],
    max_tokens: Math.round(Number.parseInt(IMAGE_MAX_WORDS) * 3),
    temperature: Number.parseFloat(IMAGE_TEMPERATURE)
  };

  const options = {
    method: "post",
    headers: {
      Authorization: "Bearer " + OPENAI_API_KEY,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  };

  log_("【REQ_2】" + JSON.stringify(data));

  const response = UrlFetchApp.fetch(url, options);
  const jsonResponse = JSON.parse(response.getContentText());
  const content = jsonResponse.choices[0].message.content.trim();

  log_("【RES_2】" + content);

  return content;
}

/**
 * OpenAIのDALL·E APIを利用して、テキストプロンプトから画像を生成します。
 *
 * @param {string} prompt - 画像生成の主要なプロンプト。
 * @param {string} imageCaption - 画像に関連するキャプション。
 * 
 * @returns {Blob} 生成された画像のBlobオブジェクト。
 *
 * 使用例:
 * const imageBlob = generateImageFromDallE_("猫", "Anime");
 * DriveApp.createFile(imageBlob);
 *
 * 注意: この関数の実行には、OPENAI_API_KEYの設定が必要です。
 */
function generateImageFromDallE_(prompt, imageCaption) {
  // 必要な情報をセットアップします
  const apiUrl = "https://api.openai.com/v1/images/generations";
  const model = "dall-e-3";
  const size = "1024x1024";
  const responseFormat = "url";

  // APIリクエストのヘッダーとペイロードを設定します
  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + OPENAI_API_KEY
    },
    payload: JSON.stringify({
      model: model,
      prompt: prompt + " " + imageCaption,
      size: size,
      quality: "hd",
      response_format: responseFormat
    })
  };

  // APIリクエストを実行し、結果を取得します
  const response = UrlFetchApp.fetch(apiUrl, options);
  const jsonResponse = JSON.parse(response.getContentText());
  const imageUrl = jsonResponse.data[0].url;
  return UrlFetchApp.fetch(imageUrl).getBlob();
}

function insertImageBlobToActiveSlide_(blob) {
  // Google スライドのアクティブなプレゼンテーションとスライドを取得
  const presentation = SlidesApp.getActivePresentation();
  const slide = presentation.getSelection().getCurrentPage();
  // 画像を挿入
  slide.insertImage(blob).setLeft(200).setTop(10);
}

function log_(message) {
  const log = `${new Date().toLocaleString()}: ${message}`;
  console.log(log);
  SpreadsheetApp.openById(SHEET_ID).getSheets()[0].appendRow([log]);
}
