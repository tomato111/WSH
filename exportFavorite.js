cplugin_exportFavorite = {
	name: 'cplugin_exportFavorite',
	commandName: 'exportFavorite',
	text: 'お気に入りをエクスポート',
	author: 'Tomato',
	description: 'StreamingPlayerのお気に入りをエクスポートします',
	onCommand: function(args) {
		var ioFunc = new function () {

		    this.createObject = function (name) {
		        return new ActiveXObject(name);
		    };

		    this.ws = this.createObject("WScript.Shell");				// WScriptのシェル
		    this.fs = this.createObject("Scripting.FileSystemObject");	// ファイル システム オブジェクト
		    this.rootDir = BASE_PATH;									// ルート ディレクトリ


    		// ファイルを読み込み、文字列を返す
    		this.readFile = function (path) {
        		if (!this.fs.FileExists(path)) return;

        		var stm = this.createObject('ADODB.Stream');
        		stm.type = 2;
        		stm.charset = "_autodetect_all";
        		stm.open();
        		try {
        			stm.loadFromFile(path);
        			var str = stm.readText(-1); //_autodetect_allでの一行ごとの取得はまともに動かない
        		} catch (e) {
	    			this.ws.Popup("ファイルが読み込めませんでした。", 0, "読み込みエラー", 48);
					return false;
        		} finally {
            		stm.Close();
            		stm = null;
				}
            	
            	return str;
    		};

    		// ファイルに書き込む
    		this.writeFile = function (text, path, CharacterCode) {
        		if (!path) return;

        		var stm = new ActiveXObject('ADODB.Stream');
        		stm.type = 2;
        		stm.charset = CharacterCode || "UTF-8";
        		stm.open();
        		stm.writeText(text);
        		try {
            		stm.saveToFile(path, 2);
        		} catch (e) {
            		this.ws.Popup("保存出来ませんでした。", 0, "書き込みエラー", 48);
            		return false;
        		} finally {
            		stm.close();
            		stm = null;
        		}

        		return true;
    		};

    		// JSON データをパースし、オブジェクトを返す
    		this.parseJSON = function (jsData) {
				return eval("("+jsData+")");
    		};

			// StreamingPlayer3のお気に入りをHTML形式に変換し、文字列を返す
			this.sp3Fav2HTML = function (data) {
				var html = "<!DOCTYPE NETSCAPE-Bookmark-file-1>\r\n<!-- This is an automatically generated file.\r\nIt will be read and overwritten.\r\nDo Not Edit! -->\r\n<TITLE>Bookmarks</TITLE>\r\n<H1>Bookmarks</H1>\r\n<DL><p>\r\n";
				var items;
				var depth = 1;
		
				if("favorites" in data)
					items = data.favorites;
				else 
					return false;
					
				createExportData(items);
				html += "</DL><p>\r\n";
				return html;
				
				function createExportData (items) {
					for (var i = 0; i<items.length; i++) {
						if (items[i].folder === true) {
							html += insertSpace(depth) + "<DT><H3 FOLDED>" + items[i].name + "</H3>\r\n" + insertSpace(depth) + "<DL><p>\r\n";
							depth++;
							arguments.callee(items[i].items);
							depth--;
							html += insertSpace(depth) + "</DL><p>\r\n";
							continue;
						}
				
						html += insertSpace(depth)
							+ "<DT><A HREF=\""
							+ items[i].videoData.uri
							+ "\" ADD_DATE=\""
							+ Date.parse(items[i].videoData.added)/1000
							+ "\" ICON_URI=\""
							+ items[i].videoData.siteIcon
							+ "\" >"
							+ items[i].videoData.title
							+ "</A>\r\n";
					}
				}

				function insertSpace (depth) {
					var space = "";
					for (var j = 0; j<depth; j++)
						space += "    ";
					return space;
				}
	
			};


		};
	
		(function () {	// main関数;
				var n, favoriteFile, jsData, SPFav, html, saved;
				var saveFile = ioFunc.ws.SpecialFolders.item("Desktop") + "\\bookmark.htm"; // 保存先
				
				n = ioFunc.ws.Popup("エクスポートしますか？(デスクトップに保存)", 0, "確認", 36);
        		if (n == 7) return;
				
				favoriteFile = ioFunc.rootDir + "\\settings\\favorites.txt"	// お気に入りファイル
				jsData = ioFunc.readFile(favoriteFile);
				if(jsData) {
					SPFav = ioFunc.parseJSON(jsData);
				}else {
					ioFunc.ws.Popup("お気に入りが読み込めなかったため処理を中止します。", 0, "処理の中止", 48);
		    		return false;
				}

				html = ioFunc.sp3Fav2HTML(SPFav);
				if (!html) {
					ioFunc.ws.Popup("変換できませんでした。", 0, "処理の中止", 48);
					return;
				}
				
				saved = ioFunc.writeFile(html, saveFile, "UTF-8");
				if (saved)
					ioFunc.ws.Popup("エクスポートが完了しました。\n" + saveFile);
					
		})();
	
	}
}
