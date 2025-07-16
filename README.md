# 丸文字自動生成ツール

## 概要

- 画面設計時にコンポーネントに番号を割り振って、リストにまとめると見やすいです。そのために作りました。
- 丸付き数字を作成するマクロ
- ExcelのShapeなので `51` 以上の問題もなし
- ３桁には未対応なのではみ出すかも
- 自動生成後には、**生成したオブジェクトを全自動で選択済みにする**機能付き

## 使い方

- 任意のフォルダにgit cloneします

```shell
git clone https://github.com/TakumiOkayasu/CircleEnclosedAlphanumericGenerater.git
cd CircleEnclosedAlphanumericGenerater
```

1. 新規Excelファイルを用意します
![NewExcelWindow](<img/01.png>)

2. 「ファイルのインポート」から `CircleCreator.bas` をインポートします
![ImportScriptFile](<img/02.png>)

3. 以下の状態になればエディタは閉じて大丈夫です。
![Import](<img/03.png>)

4. G13には始めたい数字、G14には生成したい数を入れます。自分は以下のようにしています。  
   （ただ適当に図形を作成して、そこにマクロを登録しただけです）
![MyConfig](<img/04.png>)

5. ボタンを押すと以下のように丸で囲まれた数字が生成されます（見やすさのため、全選択状態を解除していますが、生成されたときは全選択されている状態です）。
![Generate](<img/05.png>)

6. 途中から生成したい場合（例えば、１０～２０を生成したい場合）、以下のように指定するとできます。
![Generate](<img/06.png>)
