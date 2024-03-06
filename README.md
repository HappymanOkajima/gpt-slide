# GPTスライドAddOn
生成AIの力で文章や画像を生成しスライド作成をサポートする、シンプルなGoogleスライド用アドオンです。同様の機能を持つ商用アドオンは複数存在しますが、生成AIの典型的なユースケースがどのように動作するのか学びながら、自分でカスタマイズし組織に導入・運用していけるのがこのアドオンのメリットです。

## 何ができるのか
アドオンはサイドバー形式で、対象のスライドに対して3つの自動生成機能を利用できます。スライドの自動生成機能で大枠を作成し、その後、特定のテキストに対して指示を与えて深掘りしたり、マッチした画像を生成して、仕上げていくと良いでしょう。

### 1.スライドの自動生成
ChatGPTのようにプロンプトで指示を与えます。生成したいページ数とAIの創造性を指定可能です。生成されたスライドはアクティブなスライドの直後に順次追加されます。
<img width="284" alt="image" src="https://github.com/HappymanOkajima/gpt-slide/assets/6194144/17bf62ea-9002-4309-8993-873d2ed55b24">

スライドは箇条書きで作成されます。例えば次のようなプロンプトを利用することで、見栄えの良いGoogleスライドを生成することができます。これは、ChatGPTでPPTXを生成してもらうより手軽です。

> AI時代のエンジニアの働き方について。各トピックは50文字程度で、時折ユーモアを交えてください。最初のページは目次を入れてください。

もちろん、コードを生成してもらうこともできます。コードは、別のテキストボックスとして生成されるので、簡単に位置やサイズを調整できます。

> Pythonでバブルソートのサンプルコードを交え、解説してください。

<img width="733" alt="image" src="https://github.com/HappymanOkajima/gpt-slide/assets/6194144/00e430ec-30c0-4e01-81b0-174faaabe7f4">

### 2.スライド内容に基づく画像生成
表示しているスライドのテキストの内容に基づき、DALL-Eで画像を生成します。画像のスタイルを指定することもできます。画像は自動でレイアウトされませんが、スライドに埋め込まれた画像となるので、簡単に位置やサイズを調整できます。

<img width="267" alt="image" src="https://github.com/HappymanOkajima/gpt-slide/assets/6194144/bae0f1b1-8cbb-43b8-8be1-e07497465ffe">

### 3.アクティブなテキストオブジェクトに基づく文章の生成
アクティブなテキストオブジェクト（テキストボックス）の内容に基づき文章を生成します。

<img width="280" alt="image" src="https://github.com/HappymanOkajima/gpt-slide/assets/6194144/9853ed71-aff6-4dd0-825f-076681559dc5">

## インストール
### とりあえず個人で試す
手っ取り早く試すには、Googleスライドの「拡張機能」からスクリプトエディタを開き（アドオンは、[Googleドキュメントにバインドされたスクリプト](https://developers.google.com/apps-script/guides/bound?hl=ja)として動作させる必要があります）、ソースコード（code.gsとsidebar.html）をコピー＆ペーストし、後述の必須プロパティとしてAPIキーを設定することです。その後、対象のドキュメントをリロードすると、「拡張機能」メニューに「GPTスライド」が表示されます。ただしこの方法では、スクリプトがバインドされたドキュメントでのみ機能が利用できるため、あくまでも個人用・テスト用です。

<img width="637" alt="image" src="https://github.com/HappymanOkajima/gpt-slide/assets/6194144/eb153225-cbab-466f-ae3f-3fe612384726">

### 組織全体で利用する
組織全体でGoogleスライドのアドオンとして利用する場合は、GCP環境を準備し、Google Workspace Marketplace にアドオンとして登録する必要があります。[公式ドキュメント](https://developers.google.com/workspace/marketplace/how-to-publish?hl=ja)などを参考に登録・設定してください。マニフェスト（appsscript.json）も必須となりますので、このリポジトリのものを参考にしてください。

### 必須プロパティ
上記いずれの場合も、以下のスクリプトプロパティを指定する必要があります。スクリプトプロパティはGASエディタから設定してください。
- OPENAI_API_KEY（必須）
- OPENAI_MODEL（必須）
- SHEET_ID（必須：ログを記録するためのスプレッドシートのIDを指定してください）
 
※これら以外のプロパティは指定しなくてもソースコード中に指定したデフォルトで問題なく動作します。

<img width="781" alt="image" src="https://github.com/HappymanOkajima/gpt-slide/assets/6194144/02e5355b-83ae-427e-b445-4e8e2075f979">

## 関連
- [OpenAI 公式サイト](https://platform.openai.com/)。APIキーの取得はこちらから。
- [Googleドキュメント版のアドオン](https://github.com/HappymanOkajima/gpt-document)もあります。
