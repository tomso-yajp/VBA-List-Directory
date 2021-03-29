read me

module：dir_module

dir_fol(p)：<br/>
ディレクトリ内のフォルダ一覧を取得します <br/>
dir_fol("ディレクトリパス") <br/>

dir_fil(p) <br/>
ディレクトリ内のファイル一覧を取得します <br/>
dir_fil("ディレクトリパス") <br/>

dir_main： <br/>
ディレクトリ内のフォルダ、ファイル一覧を取得するメインマクロです <br/>

dir_sub(p) <br/>
ディレクトリ内のフォルダ、ファイル一覧を取得します <br/>
変数 fol,fil に格納します <br/>
dir_sub("ディレクトリパス") <br/>

new_dir： <br/>
テスト用のフォルダとファイルを作成します <br/>
作成されるディレクトリは、実行ファイルのカレントパスです <br/>

del_dir(p)： <br/>
指定のフォルダ、ファイルを削除します <br/>
del_dir("ディレクトリパス") <br/>

output_sheet(v,sname)： <br/>
取得したディレクトリの一覧をシートに追加します <br/>
output_sheet("fol,fil のどちらか","シート名") <br/>

sheet_list()： <br/>
全てのシート名を取得します <br/>
戻り値は、"シート名1, シート名2, ..." <br/>



