#### Gemini Excel Quiery GEQについて
- シート名で顧客とカテゴリー(Macro/M-MIMO)を区別しているExcelを仮定
- FakePrograms.xlsxで それらを仮に A, B, C, Dと簡素化してシートを作成
- このExcelシートには非常に単純なプログラムの周波数や MIMO数などのランダムな表を作成
- これらを読み取って Gemini AIにAPIで指示とともに投げて返答を受けとる
- geq3.pyでほぼ問題なく動作していることがわかったので app.pyにcopy

