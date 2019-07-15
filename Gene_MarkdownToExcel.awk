# Gene_MarkdownToExcel.awk
# gawk.exe -f Gene_MarkdownToExcel.awk define.txt

# https://ryuta46.com/255
# Markdown で書いた試験仕様書を Excel に変換するツールを作った

# の、kotlinからpowershellとかgoのexcelizeへの移植を兼ねたもの

# ------------------------------------------------------------------------------------------------------------------------------------

# 現在使っているのは、h1、h2、h3、h4 要素と 順序付きの項目、チェックボックスになります。それぞれの役割としては
# h1 は試験のカテゴリ。コンバートする時はこの要素が Excel の 1シートになります。
# h2、h3、h4 はそれぞれ大項目、中項目、小項目
# 順序付きの項目は試験手順
# チェックボックスは確認内容
# という感じです。

# ------------------------------------------------------------------------------------------------------------------------------------

BEGIN{
	GeneTime = strftime("%Y/%m/%d %H:%M:%S", systime());
	print "# "GeneTime > "GeneExcelize.ps1";
	# Excel の起動
	print "$excel = New-Object -ComObject Excel.Application" > "GeneExcelize.ps1";
	# 不可視化
	print "$excel.Visible = $false" > "GeneExcelize.ps1";
	# アラートを無効に
	print "$excel.DisplayAlerts = $False" > "GeneExcelize.ps1";
	
}

BEGINFILE{
	ElementCount = 0;
	RowNum = 2;
	Cas4_Flg = "false";
}

# <!--.*?-->と空行は飛ばす
/<!--.*?-->/{
	next;
}

/./{
	switch(ElementCount){
		case 0:
			if(substr($0,1,2) == "# "){
				Out_H1();
				ElementCount++;
				next;
			} else {
				print "構文エラー : h1要素がないお・・・" > "con";
				exit 99;
			}
			break;
		case 1:
			if(substr($0,1,3) == "## "){
				Out_H2();
				ElementCount++;
				next;
			} else {
				print "構文エラー : h2要素がないお・・・" > "con";
				exit 99;
			}
			break;
		case 2:
			if(substr($0,1,4) == "### "){
				Out_H3();
				ElementCount++;
				next;
			} else {
				print "構文エラー : h3要素がないお・・・" > "con";
				exit 99;
			}
			break;
		case 3:
			if(substr($0,1,5) == "#### "){
				Out_H4();
				ElementCount++;
				next;
			} else {
				print "構文エラー : h4要素がないお・・・" > "con";
				exit 99;
			}
			break;
		case 4:
			# 確認手順
			# 1回も確認手順を処理していないのに、「* [ ] 」やh1-h4要素を出してきた場合
			if(((Cas4_Flg == "false") && (substr($0,1,1) == "#")) || ((Cas4_Flg == "false") && (substr($0,1,6) == "* [ ] "))){
				print "構文エラー : 確認手順が1つもないお・・・" > "con";
				exit 99;
			} else {
				Cas4_Flg = "true";
			}
			# とりあえず、1回は確認手順を処理したが、「* [ ] 」を出してきた場合
			if((substr($0,1,6) == "* [ ] ")){
				if(length(escs) == 1){
					esc = escs[i];
					gsub("\"","`\"",esc);
					print "$sheet.Range(\"E"RowNum"\") = \""esc"\"" > "GeneExcelize.ps1";
					delete escs;
				} else {
					for(i in escs){
						esc = escs[i];
						gsub("\"","`\"",esc);
						if(i == 1){
							print "$sheet.Range(\"E"RowNum"\") = \""esc > "GeneExcelize.ps1";
						} else if(i == length(escs)){
							print esc"\"" > "GeneExcelize.ps1";
						} else {
							print esc > "GeneExcelize.ps1";
						}
						esc = "";
					}
					delete escs;
					ElementCount++;
					escs_cnt = 1;
				}
			} else {
				# とりあえず、1回は確認手順を処理した
				escs[escs_cnt] = $0;
				escs_cnt++;
				next;
			}
			break;
	}
}

# 確認項目
/./{
	if(ElementCount == 5){
		# h1-h4要素が出てくるまで、配列に格納する
		if(substr($0,1,1) == "#"){
			if(length(escs) == 1){
				esc = escs[i];
				gsub("\"","`\"",esc);
				print "$sheet.Range(\"F"RowNum"\") = \""esc"\"" > "GeneExcelize.ps1";
				delete escs;
			} else {
				for(i in escs){
					esc = escs[i];
					gsub("\"","`\"",esc);
					if(i == 1){
						print "$sheet.Range(\"F"RowNum"\") = \""esc > "GeneExcelize.ps1";
					} else if(i == length(escs)){
						print esc"\"" > "GeneExcelize.ps1";
					} else {
						print esc > "GeneExcelize.ps1";
					}
					esc = "";
				}
				delete escs;
				escs_cnt = 1;
			}
		} else if((substr($0,1,6) == "* [ ] ")){
			esc = "・"substr($0,7);
			escs[escs_cnt] = esc;
			escs_cnt++;
			esc = "";
			next;
		} else {
			print "構文エラー" > "con";
			exit 99;
		}
	}
}

/./{
	RowNum++;
	if(substr($0,1,5) == "#### "){
		Out_H4();
		ElementCount = 4;
		next;
	} else if(substr($0,1,4) == "### "){
		Out_H3();
		ElementCount = 3;
		next;
	} else if(substr($0,1,3) == "## "){
		print "$tableRange = $sheet.Range($sheet.Cells.Range(\"C"Out_H2_StartRow"\"), $sheet.Cells.Range(\"C"RowNum"\"))" > "GeneExcelize.ps1";
		Out_H2_StartRow = RowNum + 1;
		print "$tableRange.MergeCells = $true" > "GeneExcelize.ps1";
		Out_H2();
		ElementCount = 2;
		next;
	} else if(substr($0,1,2) == "# "){
		print "$tableRange = $sheet.Range($sheet.Cells.Range(\"B"Out_H1_StartRow"\"), $sheet.Cells.Range(\"B"RowNum"\"))" > "GeneExcelize.ps1";
		Out_H1_StartRow = RowNum + 1;
		print "$tableRange.MergeCells = $true" > "GeneExcelize.ps1";
		Ender();
		Out_H1();
		ElementCount = 1;
		next;
	} else {
		print "構文エラー : h4要素がないお・・・" > "con";
		exit 99;
	}
}

ENDFILE{
	if(length(escs) == 1){
		esc = escs[i];
		gsub("\"","`\"",esc);
		print "$sheet.Range(\"F"RowNum"\") = \""esc"\"" > "GeneExcelize.ps1";
		delete escs;
	}else {
		for(i in escs){
			esc = escs[i];
			gsub("\"","`\"",esc);
			if(i == 1){
				print "$sheet.Range(\"F"RowNum"\") = \""esc > "GeneExcelize.ps1";
			} else if(i == length(escs)){
				print esc"\"" > "GeneExcelize.ps1";
			} else {
				print esc > "GeneExcelize.ps1";
			}
			esc = "";
		}
		delete escs;
		escs_cnt = 1;
	}
	Ender();
}

# h1要素
function Out_H1(){
	RowNum = 2;
	# シート名（試験のカテゴリ、らしいお・・・）
	print "$book = $excel.Workbooks.Add()"  > "GeneExcelize.ps1";
	print "$sheet = $book.ActiveSheet" > "GeneExcelize.ps1";
	print "$sheet.Name = \""substr($0,3)"\"" > "GeneExcelize.ps1";
	# 見出しをこの時点で作っておくか。
	print "$sheet.Range(\"B"RowNum"\") = \"大項目\"" > "GeneExcelize.ps1";
	print "$sheet.Range(\"C"RowNum"\") = \"中項目\"" > "GeneExcelize.ps1";
	print "$sheet.Range(\"D"RowNum"\") = \"小項目\"" > "GeneExcelize.ps1";
	print "$sheet.Range(\"E"RowNum"\") = \"確認手順\"" > "GeneExcelize.ps1";
	print "$sheet.Range(\"F"RowNum"\") = \"確認項目\"" > "GeneExcelize.ps1";
	print "$sheet.Range(\"G"RowNum"\") = \"結果\"" > "GeneExcelize.ps1";
	print "$sheet.Range(\"H"RowNum"\") = \"試験日\"" > "GeneExcelize.ps1";
	print "$sheet.Range(\"I"RowNum"\") = \"試験者\"" > "GeneExcelize.ps1";
	print "$sheet.Range(\"J"RowNum"\") = \"備考\"" > "GeneExcelize.ps1";
	print "$tableRange = $sheet.Range($sheet.Cells.Range(\"B"RowNum"\"), $sheet.Cells.Range(\"J"RowNum"\"))" > "GeneExcelize.ps1";
	RowNum++;
	Out_H1_StartRow = RowNum;
	print "$tableRange.interior.ColorIndex = 37" > "GeneExcelize.ps1";
}

# h2要素
function Out_H2(){
	# 大項目
	print "$sheet.Range(\"B"RowNum"\") = \""substr($0,4)"\"" > "GeneExcelize.ps1";
	Out_H2_StartRow = RowNum;
}

# h3要素
function Out_H3(){
	# 中項目
	print "$sheet.Range(\"C"RowNum"\") = \""substr($0,5)"\"" > "GeneExcelize.ps1";
}

# h4要素
function Out_H4(){
	# 小項目
	print "$sheet.Range(\"D"RowNum"\") = \""substr($0,6)"\"" > "GeneExcelize.ps1";
	esc = "";
	escs_cnt = 1;
}

function Ender(){
	print "$tableRange = $sheet.Range($sheet.Cells.Range(\"C"Out_H2_StartRow"\"), $sheet.Cells.Range(\"C"RowNum"\"))" > "GeneExcelize.ps1";
	print "$tableRange.MergeCells = $true" > "GeneExcelize.ps1";
	print "$tableRange = $sheet.Range($sheet.Cells.Range(\"B"Out_H1_StartRow"\"), $sheet.Cells.Range(\"B"RowNum"\"))" > "GeneExcelize.ps1";
	print "$tableRange.MergeCells = $true" > "GeneExcelize.ps1";
	print "$tableRange = $sheet.Range($sheet.Cells.Range(\"B2\"), $sheet.Cells.Range(\"J"RowNum"\"))" > "GeneExcelize.ps1";
	print "$tableRange.Borders.LineStyle = $True" > "GeneExcelize.ps1";
	print "$tableRange.Columns.AutoFit() | Out-Null" > "GeneExcelize.ps1";
	print "$tableRange = $sheet.Range($sheet.Cells.Range(\"E3\"), $sheet.Cells.Range(\"E10000\"))" > "GeneExcelize.ps1";
	print "$tableRange.Columns.AutoFit() | Out-Null" > "GeneExcelize.ps1";
	print "$tableRange.HorizontalAlignment = -4131" > "GeneExcelize.ps1";
	print "$tableRange = $sheet.Range($sheet.Cells.Range(\"F3\"), $sheet.Cells.Range(\"F10000\"))" > "GeneExcelize.ps1";
	print "$tableRange.Columns.AutoFit() | Out-Null" > "GeneExcelize.ps1";
	print "$tableRange.HorizontalAlignment = -4131" > "GeneExcelize.ps1";
	print "$sheet.Cells.Item(1,1).Select() | Out-Null" > "GeneExcelize.ps1";
	print "$excel.ActiveWindow.Zoom = 100" > "GeneExcelize.ps1";
}

END{
	print "$excel.Worksheets.Item(1).Activate()" > "GeneExcelize.ps1";
	GeneTime = strftime("%Y%m%d_%H%M%S", systime());
	cmd = "cd";
	while(cmd | getline esc){
		break;
	}
	close(cmd);
	print "$book.SaveAs(\""esc"\\Specifications_"GeneTime".xlsx\")" > "GeneExcelize.ps1";
	print "$excel.Quit()" > "GeneExcelize.ps1";
	print "$excel = $Null" > "GeneExcelize.ps1";
	print "[GC]::collect()" > "GeneExcelize.ps1";
}

