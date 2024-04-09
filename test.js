function test_createSlidesInCurrentPresentation_1() {
  createSlidesInCurrentPresentation('日本の気候の特徴について。',3,"mid");
}
function test_createSlidesInCurrentPresentation_2() {
  createSlidesInCurrentPresentation('テーマは徳川吉宗の生涯。JavaScriptで表示するコードのサンプルも添えて。',3,"high");
}
function test_createSlidesInCurrentPresentation_3() {
  createSlidesInCurrentPresentation('トピックは、コサイン類似度の考え方。JavaScriptのサンプルソースもください。',5,"low");
}
function test_generateImageInCurrentSlide() {
  generateImageInCurrentSlide("4k, mono");
}
function test_generateTextInCurrentShape() {
  generateTextInCurrentElement("50文字に要約して");
}
