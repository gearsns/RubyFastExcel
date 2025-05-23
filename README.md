# RubyFastExcel
RubyでEXCELファイルを出力するサンプル

## 概要

「[Rubyで最速のxlsx生成プログラムを作ってみた](https://rooter.jp/data-format/fastest-xlsx-generatior-ruby/)」で、xlsxを作成っていうのがあったので折角なのでテーブル機能と標準出力に出力する機能も実装してみた

## 使い方

require
``` ruby
require_relative 'fast_xlsx.rb'
```

ファイルに書き出し
``` ruby
xlsx = FastXLSX.new
1.upto(10) do |i|
  xlsx.add_row([i,"col1value#{i}", "col2value#{i}", "col3value#{i}", Time.now])
end
xlsx.save("test.xlsx")
```

標準出力に書き出し
``` ruby
xlsx = FastXLSX.new
1.upto(10) do |i|
  xlsx.add_row([i,"col1value#{i}", "col2value#{i}", "col3value#{i}", Time.now])
end
xlsx.save()
```

テーブル機能ありでファイルに書き出し
``` ruby
xlsx = FastXLSX.new
xlsx.set_header(%w(a a b c d))
1.upto(10) do |i|
  xlsx.add_row([i,"col1value#{i}", "col2value#{i}", "col3value#{i}", Time.now])
end
xlsx.save("test2.xlsx")
```

テーブル機能ありで標準出力に書き出し
``` ruby
xlsx = FastXLSX.new
xlsx.set_header(%w(a a b c d))
1.upto(10) do |i|
  xlsx.add_row([i,"col1value#{i}", "col2value#{i}", "col3value#{i}", Time.now])
end
xlsx.save()
```

