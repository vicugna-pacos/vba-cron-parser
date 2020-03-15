# Cron parser in VBA
VBAでcronのスケジュール設定(`0 0 12 ? * WED`とかのやつ)をパースする。
cronといいつつBizRoboのスケジュール管理をしたかったので、cronの文法はBizRoboに準拠している。

参考：[Cron(クーロン)形式のスケジュール – BizRobo! ナレッジベース](https://knowledge.bizrobo.com/hc/ja/articles/360034654472-Cron-%E3%82%AF%E3%83%BC%E3%83%AD%E3%83%B3-%E5%BD%A2%E5%BC%8F%E3%81%AE%E3%82%B9%E3%82%B1%E3%82%B8%E3%83%A5%E3%83%BC%E3%83%AB)

# 曜日
cronの場合、0=日曜日、1=月曜日...となっているが、BizRoboのMCは1=日曜日、2=月曜日...となっている。
このスクリプトを作ったきっかけはBizRoboなので、後者を採用している。
曜日は「SUN」などの名前で指定することもできるので、こちらを使用すればズレは生じない…はず。


# VBAのソース管理
https://github.com/vbaidiot/ariawase

VBAライブラリ「Ariawase」に含まれている「vbac」を利用させていただいている。


