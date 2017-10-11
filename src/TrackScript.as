package  {

	import com.childoftv.xlsxreader.Worksheet;
	import com.childoftv.xlsxreader.XLSXLoader;
	import fl.controls.Button;
	import fl.controls.TextInput;
	import flash.display.NativeWindow;
	import flash.display.Sprite;
	import flash.display.StageAlign;
	import flash.display.StageScaleMode;
	import flash.errors.IOError;
	import flash.events.ErrorEvent;
	import flash.events.Event;
	import flash.events.KeyboardEvent;
	import flash.events.MouseEvent;
	import flash.events.NativeWindowBoundsEvent;
	import flash.events.ProgressEvent;
	import flash.events.UncaughtErrorEvent;
	import flash.filesystem.File;
	import flash.filesystem.FileMode;
	import flash.filesystem.FileStream;
	import flash.globalization.DateTimeFormatter;
	import flash.net.FileFilter;
	import flash.system.Capabilities;
	import flash.text.TextFormat;
	import flash.ui.Keyboard;
	import flash.utils.ByteArray;

	/**
	 * ...
	 * @author Tim Yusupov
	 */
	public class TrackScript extends Sprite {

		private var main:Main;
		private var ui:UI;
		private var win:NativeWindow;

		private const COLOR_BAD:String = "#CC171C"; // Red
		private const COLOR_SUCCESS:String = "#189510"; // Green
		private const COLOR_WARN:String = "#CB5815"; // Orange

		private var xlFilePathError:Boolean = false;
		private var xmlFilePathError:Boolean = false;
		private var devFlag:Boolean = false;

		public function TrackScript():void {
			stage ? init() : addEventListener(Event.ADDED_TO_STAGE, init);
		}

		private function init(e:Event = null):void {

			removeEventListener(Event.ADDED_TO_STAGE, init);

			// Entry Point
			main = Main.ins;
			ui = new UI();
			addChild(ui);

			win = stage.nativeWindow;
			win.title = "TrackScript";
			ui.tfVer.text = "v" + main.version;

			var winPosStr:String = main.settings.getKey(Settings.winPos);
			var reResult:Array = winPosStr.match(/(\d+):(\d+)/);
			win.x = Number(reResult[1]);
			win.y = Number(reResult[2]);

			win.addEventListener(NativeWindowBoundsEvent.MOVE, function(e:NativeWindowBoundsEvent):void {
				main.settings.setKey(Settings.winPos, win.x + ":" + win.y);
			});

			win.addEventListener(Event.CLOSING, function(e:Event):void {
				e.preventDefault();
				main.exitApp();
			});

			win.activate();

			// Stage Settings
			stage.scaleMode = StageScaleMode.NO_SCALE;
			stage.align = StageAlign.TOP;

			// Keyboard
			stage.addEventListener(KeyboardEvent.KEY_DOWN, keyDown);

			// UI
			// Styles
			var defTextFormat:TextFormat = new TextFormat("Tahoma", 12);
			var btnTextFormat:TextFormat = new TextFormat(defTextFormat.font, 16);
			var tfTextFormat:TextFormat = new TextFormat(defTextFormat.font, defTextFormat.size);
			tfTextFormat.leftMargin = tfTextFormat.rightMargin = 3;

			ui.tfXlFile.setStyle("textFormat", tfTextFormat);
			ui.tfXlFile.setStyle("disabledTextFormat", tfTextFormat);
			ui.tfXmlFile.setStyle("textFormat", tfTextFormat);
			ui.tfXmlFile.setStyle("disabledTextFormat", tfTextFormat);
			ui.taOutput.setStyle("textFormat", defTextFormat);
			ui.taOutput.setStyle("disabledTextFormat", defTextFormat);
			ui.btnStart.setStyle("textFormat", btnTextFormat);
			ui.btnStart.setStyle("disabledTextFormat", btnTextFormat);

			ui.tfXlFile.text = main.settings.getKey(Settings.sourceExcelFile);
			ui.tfXmlFile.text = main.settings.getKey(Settings.trackCheckerDataFile);
			ui.tfXlFile.addEventListener("change", onTfChange);
			ui.tfXmlFile.addEventListener("change", onTfChange);

			ui.taOutput.editable = false;

			ui.btnXlDialog.addEventListener(MouseEvent.CLICK, btnDialogClick);
			ui.btnXmlDialog.addEventListener(MouseEvent.CLICK, btnDialogClick);
			ui.btnStart.addEventListener(MouseEvent.CLICK, btnStartClick);
			ui.btnStart.label = "З А П У С К";

			xlFile = new File();
			xmlFile = new File();
			xlFile.addEventListener(Event.SELECT, onFileSelect);
			xmlFile.addEventListener(Event.SELECT, onFileSelect);

			loaderInfo.uncaughtErrorEvents.addEventListener(UncaughtErrorEvent.UNCAUGHT_ERROR, onUncaughtError);

			// Check services file
			var srvsFile:File = File.applicationStorageDirectory.resolvePath("services.txt");
			if (!srvsFile.exists) {
				fst = new FileStream();
				fst.open(srvsFile, FileMode.WRITE);
				fst.writeUTFBytes(srvsFileDefaultContent);
				fst.close();
			}

			// Load services file
			fst = new FileStream();
			fst.open(srvsFile, FileMode.READ);
			var srvsFileString:String = fst.readUTFBytes(fst.bytesAvailable);
			fst.close();

			srvsObj = parseServicesFile(srvsFileString);

			// Check dev marker-file
			if (File.applicationStorageDirectory.resolvePath("dev").exists)
				devFlag = true;

		}

		private function onUncaughtError(e:UncaughtErrorEvent):void {
            if (e.error is Error)
            {
                var error:Error = e.error as Error;
				outputLogLine("Ошибка: " + error.message, COLOR_BAD);
            }
		}

		private function btnStartClick(e:MouseEvent):void {
			start();
		}

		private function keyDown(e:KeyboardEvent):void {
			// ESC
			if (e.keyCode == Keyboard.ESCAPE) {

				main.exitApp();

			}

			else

			// ENTER
			if (e.keyCode == Keyboard.ENTER) {
				start();
			}
		}

		private function onTfChange(e:Event):void {
			var tf:TextInput = e.target as TextInput;

			if (tf == ui.tfXlFile) {

				main.settings.setKey(Settings.sourceExcelFile, ui.tfXlFile.text);

				try {
					xlFile.nativePath = ui.tfXlFile.text;
					ui.tfXlFile.htmlText = colorText("#000000", ui.tfXlFile.text);
					xlFilePathError = false;
				}

				catch (e:Error) {
					ui.tfXlFile.htmlText = colorText(COLOR_BAD, ui.tfXlFile.text);
					xlFilePathError = true;
				}

			}
			else
			if (tf == ui.tfXmlFile) {

				main.settings.setKey(Settings.trackCheckerDataFile, ui.tfXmlFile.text);

				try {
					xmlFile.nativePath = ui.tfXmlFile.text;
					ui.tfXmlFile.htmlText = colorText("#000000", ui.tfXmlFile.text);
					xmlFilePathError = false;
				}

				catch (e:Error) {
					ui.tfXmlFile.htmlText = colorText(COLOR_BAD, ui.tfXmlFile.text);
					xmlFilePathError = true;
				}

			}

			validateInputs();
		}

		private function onFileSelect(e:Event):void {
			var file:File = e.target as File;

			if (file == xlFile) {
				ui.tfXlFile.text = file.nativePath;
				ui.tfXlFile.dispatchEvent(new Event("change"));
			}
			else
			if (file == xmlFile) {
				ui.tfXmlFile.text = file.nativePath;
				ui.tfXmlFile.dispatchEvent(new Event("change"));
			}
		}

		private function btnDialogClick(e:MouseEvent):void {
			var btn:Button = e.target as Button;

			if (btn == ui.btnXlDialog) {
				xlFile.browseForOpen("Файл Excel 2007 с информацией о треках", [new FileFilter("Excel 2007", "*.xlsx")]);
			}
			else
			if (btn == ui.btnXmlDialog) {
				xmlFile.browseForOpen("Файл \"data.xml\" от TrackChecker", [new FileFilter("Файл XML", "*.xml")]);
			}
		}

		private function validateInputs():Boolean {

			// Empty paths
			if (ui.tfXlFile.text == "" || ui.tfXmlFile.text == "") {
				return false;
			}

			// Not that extension
			if (ui.tfXlFile.text.search(/.xlsx$/) == -1) {
				ui.tfXlFile.htmlText = colorText(COLOR_BAD, ui.tfXlFile.text);
				return false;
			} else {
				ui.tfXlFile.htmlText = colorText("#000000", ui.tfXlFile.text);
			}

			if (ui.tfXmlFile.text.search(/.xml$/) == -1) {
				ui.tfXmlFile.htmlText = colorText(COLOR_BAD, ui.tfXmlFile.text);
				return false;
			} else {
				ui.tfXmlFile.htmlText = colorText("#000000", ui.tfXmlFile.text);
			}

			// Not existing
			try {

				if (!xlFile.exists) {
					ui.tfXlFile.htmlText = colorText(COLOR_BAD, ui.tfXlFile.text);
					return false;
				} else {
					ui.tfXlFile.htmlText = colorText("#000000", ui.tfXlFile.text);
				}

			}

			catch (e:Error) {
				trace("АШИБКА");
			}


			try {

				if (!xmlFile.exists) {
					ui.tfXmlFile.htmlText = colorText(COLOR_BAD, ui.tfXmlFile.text);
					return false;
				} else {
					ui.tfXmlFile.htmlText = colorText("#000000", ui.tfXmlFile.text);
				}

			}

			catch (e:Error) {
				trace("АШИБКА");
			}

			return true;

		}

		private function generalValidate():Boolean {
			var status:Boolean;

			if (xlFilePathError || xmlFilePathError) {
				return false;
			}

			status = validateInputs();

			return status;
		}

		private function outputLogLine(tx:String, color:String = null):void {

			if (tx == "=") {
				tx = "=============================";
			}

			ui.taOutput.htmlText += (color != null ? colorText(color, tx) : tx) + "\n";
			ui.taOutput.verticalScrollPosition = ui.taOutput.maxVerticalScrollPosition;

		}

		private var fst:FileStream;
		private var xlLoader:XLSXLoader;
		private var xlSheet:Worksheet;
		private var xlFile:File;
		private var xmlFile:File;
		private var dataXml:XML;
		private var xmlString:String;

		private var maxIDDefault:uint;
		private var maxID:uint;
		private var currentDate:String;

		private var tracksCount:uint;
		private var existingGroupsCount:uint;
		private var existingTracksCount:uint;

		private var active:Boolean;
		private var row:int;
		private var groupMode:Boolean;
		private var groups:XMLList;
		private var currentGroup:XML;
		private var emptyLinesCount:int;
		private var wholeLineIsEmpty:Boolean;

		private var srvsObj:Object = {};

		private function start():void {

			if (!generalValidate()) {
				outputLogLine("Имеются ошибки. Запуск невозможен.", COLOR_BAD);
				return;
			}

			ui.taOutput.text = "";

			// Load excel file
			xlLoader = new XLSXLoader();
			xlLoader.addEventListener(Event.COMPLETE, start2);
			xlLoader.load(ui.tfXlFile.text);
			outputLogLine("Загрузка файла Excel");

		}

		private function start2(e:Event):void {

			xlLoader.removeEventListener(Event.COMPLETE, start2);

			// XML File Reading Start
			xmlFile.nativePath = ui.tfXmlFile.text;
			fst = new FileStream();

			fst.addEventListener(ProgressEvent.PROGRESS, onXMLFileLoadingProgress);
			fst.addEventListener(Event.COMPLETE, start3);
			fst.openAsync(xmlFile, FileMode.READ);
			outputLogLine("Загрузка файла \"data.xml\"");

		}

		private function start3(e:Event):void {

			// XML
			trace("XML Data File: reading done");
			fst.removeEventListener(Event.COMPLETE, start3);
			trace("Bytes Available " + fst.bytesAvailable);
			xmlString = fst.readUTFBytes(fst.bytesAvailable)
			fst.close();

			trace("XML String Length:", xmlString.length);

			dataXml = new XML(xmlString);

			var rootNodesCount:uint = dataXml.groups.children().length();
			trace("Nodes Count:", rootNodesCount);

			if (dataXml.@maxid == "" || dataXml.@maxid == null) {
				outputLogLine("Неверный формат \"data.xml\"", COLOR_BAD);
				return;
			}

			maxID = maxIDDefault = uint(dataXml.@maxid);
			trace("MAX ID:", dataXml.@maxid);

			outputLogLine("MaxID: " + String(maxID) + "; " + "Элементов в корневой группе: " + String(rootNodesCount));

			// EXCEL
			xlSheet = xlLoader.worksheet("Sheet1");

			if (xlSheet.getCellValue("A1").search(/Parcel List/i) == -1) {
				outputLogLine("Неверный формат Excel", COLOR_BAD);
				return;
			}

			//outputLogLine("=");

			/**
			 * Parsing Excel
			 * ================================================================================
			 */
			tracksCount = existingGroupsCount = existingTracksCount = 0;

			row = 3;
			groupMode = false;
			emptyLinesCount = 0;
			wholeLineIsEmpty = false;

			groups = new XMLList(<track-script-output></track-script-output>);

			var groupColVal:String;
			var trackColVal:String;
			var adrColVal:String;
			//var nameColVal:String;
			//var cntColVal:String;

			var trackCol:String;
			var trackColumns:Array = ["J"];

			var groupName:String;
			var name:String;
			var track:String;
			var country:String;

			// Find column with tracks
			var colVal:String;
			for each (var trco:String in trackColumns) {
				colVal = xlSheet.getCellValue(trco + row);
				colVal = trimSpaces(colVal);
				if (colVal.search(/^[a-zA-Z]{2}\d{9}(CN|cn|Cn|cN)?$/) != -1) {
					trackCol = trco;
					break;
				}
			}

			if (trackCol == null) {
				outputLogLine("Ошибка: не найдена колонка с треками", COLOR_BAD);
				return;
			}

			outputLogLine("Колонка с треками: " + trackCol);

			currentDate = printDate();
			active = true;

			while (active) {

				// Reset some variables
				wholeLineIsEmpty = false;

				groupColVal = xlSheet.getCellValue("C" + row);
				trackColVal = xlSheet.getCellValue(trackCol + row);
				adrColVal = xlSheet.getCellValue("H" + row);
				//nameColVal = xlSheet.getCellValue("E" + row);
				//cntColVal = xlSheet.getCellValue("J" + row);

				if (groupColVal == "" && trackColVal == "" && adrColVal == "") {
					wholeLineIsEmpty = true;
				}

				if (adrColVal == "" || adrColVal == null) {

					!wholeLineIsEmpty && outputLogLine("Пустой адрес на строке " + row, COLOR_WARN);

				} else {

					var adrObj:Object = parseAddressCell(adrColVal);
					name = adrObj.name;
					country = adrObj.country;

				}

				if (trackColVal == "") {
					!wholeLineIsEmpty && outputLogLine("Пустой трек на строке " + row, COLOR_WARN);
				}

				if (trackColVal != "" && emptyLinesCount == 1) {
					emptyLinesCount = 0;
				}

				row++;

				if (emptyLinesCount > 1) {
					active = false;
					break;
				}

				if (trackColVal == "") {
					emptyLinesCount++;
					continue;
				}

				if (groupColVal != "") {
					groupMode = true;
					groupName = trimSpaces(groupColVal);

					currentGroup = <groups></groups>;
					currentGroup.@id = ++maxID;
					currentGroup.@desc = groupName;
					currentGroup.@crdt = currentDate;

					groups.appendChild(currentGroup);
				}

				track = trimSpaces(trackColVal);
				//name = trimSpaces(nameColVal);
				//country = trimSpaces(cntColVal);

				var xmlTrack:XML = new XML(<track></track>);
				var xmlTrackServs:XML = new XML(<servs></servs>);
				var xmlServ:XML;

				xmlTrack.@id = ++maxID;
				xmlTrack.@desc = name;
				xmlTrack.@crdt = currentDate;
				xmlTrack.@track = track;

				xmlTrackServs.@id = ++maxID;
				xmlTrackServs.@crdt = currentDate;

				var servAliases:Array = ["china", "china_alt", "china_ems"];
				var servAlias:String;

				for each (servAlias in servAliases) {
					xmlServ = new XML(<serv/>);
					xmlServ.@id = ++maxID;
					xmlServ.@crdt = currentDate;
					xmlServ.@serv = servAlias;
					xmlServ.@selected = 1;
					xmlTrackServs.appendChild(xmlServ);
				}

				var specialServs:Array = getCntServices(country);

				if (specialServs != null) {

					for each (servAlias in specialServs) {
						xmlServ = new XML(<serv/>);
						xmlServ.@id = ++maxID;
						xmlServ.@crdt = currentDate;
						xmlServ.@serv = servAlias;
						xmlServ.@selected = 1;
						xmlTrackServs.appendChild(xmlServ);
					}

				}

				// Special event "Added". Appended to every track
				var xmlEventAdded:XML = new XML(<event/>);
				xmlEventAdded.@id = ++maxID;
				xmlEventAdded.@crdt = currentDate;
				xmlEventAdded.@desc = "Added";
				xmlEventAdded.@udt = printDate(true);

				xmlTrack.appendChild(xmlTrackServs);
				xmlTrack.appendChild(xmlEventAdded);
				currentGroup.appendChild(xmlTrack);
				tracksCount++;

			}

			// Write result of script work to output.xml
			var d:Date = new Date();
			var dtf:DateTimeFormatter = new DateTimeFormatter("ru-RU");
			var dstr:String;
			dtf.setDateTimePattern("dd.MM.yyyy HH:mm:ss");
			dstr = dtf.format(d);

			groups.@date = dstr;
			groups.@version = main.version;
			groups.@excel = ui.tfXlFile.text;

			XML.prettyPrinting = true;
			XML.prettyIndent = 4;
			var outputFile:File = File.applicationStorageDirectory.resolvePath("output.xml");
			fst.openAsync(outputFile, FileMode.WRITE);
			fst.writeUTFBytes(groups.toXMLString());
			fst.close();

			/**
			 * Filling source XML (dataXml) with generated groups
			 * ================================================================================
			 */
			// Iterate through our groups
			for each (var grp:XML in groups.groups) {

				trace("Group:", grp.@desc);

				// Check if this group already exists
				// If exists
				var x:* = dataXml.groups.(@id == 0)..groups.(@desc == grp.@desc);
				if (x.length() > 0) {

					// Compare tracks in existing group to our current group
					for each (var tr:XML in groups.groups.(@desc == grp.@desc).track) {
						trace("Track:", tr.@track);

						// Track Existence Check
						// If found duplicate > skip this track
						x = dataXml.groups.(@id == 0)..groups.(@desc == grp.@desc).track.(@desc == tr.@desc && @track == tr.@track);
						if (x.length() > 0) {
							main.logRed("Track Duplicate Found: " + tr.@desc + " " + tr.@track);
							existingTracksCount++;
							continue;
						}

						// If not found > add this track to existing group
						else {

							for each (var subGroup:XML in dataXml.groups.(@id == 0)..groups.(@desc == grp.@desc)) {
								subGroup.appendChild(tr);
							}
							//dataXml.groups.(@id == 0)..groups.(@desc == grp.@desc)[0].appendChild(tr);

						}
					}

					// Stats
					existingGroupsCount++;

					// Skip to our next group
					continue;

				}

				// Add new groups to root group in Source XML
				dataXml.groups.(@id == 0).appendChild(grp);

			}

			/**
			 * Show Stats
			 * ================================================================================
			 */
			outputLogLine("Найдено групп: " + groups.children().length());
			outputLogLine("Обработано треков: " + tracksCount);
			if (existingGroupsCount > 0) outputLogLine("Группы, которые уже есть в программе: " + existingGroupsCount);
			if (existingTracksCount > 0) outputLogLine("Найденные дубли треков: " + existingTracksCount);

			if (tracksCount == existingTracksCount && tracksCount > 0) {
				outputLogLine("Похоже кто-то прогнал меня два раза", COLOR_BAD);
			}

			if (Capabilities.isDebugger || devFlag) {outputLogLine("Готово", COLOR_SUCCESS); return;}

			/**
			 * Write Back to XML File
			 * ================================================================================
			 */
			XML.prettyPrinting = false;
			var outputStr:String = dataXml.toXMLString();

			fst.openAsync(xmlFile, FileMode.WRITE);
			fst.writeUTFBytes(outputStr);
			fst.close();

			outputLogLine("Готово", COLOR_SUCCESS);

		}

		private function onXMLFileLoadingProgress(e:ProgressEvent):void {
			trace("Loading XML File. Progress:",
				e.bytesLoaded + " / " + e.bytesTotal,
				Math.floor(e.bytesLoaded / e.bytesTotal * 100) + "%"
			);
		}

		private function parseAddressCell(adrCellVal:String):Object {

			var ctrlCharPattern:RegExp = /(\r|\n|\r\n)/;
			var name:String;
			var country:String;

			var adrColVal:String = trimSpaces(adrCellVal);

			// Check: empty or one line
			if (adrColVal.length < 1 || adrColVal.search(ctrlCharPattern) == -1) {
				// return
			}

			var adrLines:Array;
			var linesTemp:Array = [];

			// Разделить по строкам
			adrLines = adrColVal.split(ctrlCharPattern);

			var i:int;

			// Отчистить от управляющих символов
			for (i = 0; i < adrLines.length; i++) {
				if ((adrLines[i] as String).search(ctrlCharPattern) == -1) {
					linesTemp.push(adrLines[i]);
				}
			}

			adrLines = linesTemp;
			linesTemp = [];

			// Отчистить от пустых символов
			for (i = 0; i < adrLines.length; i++) {
				if ((adrLines[i] as String).length != 0 || (adrLines[i] as String) != "") {
					linesTemp.push(adrLines[i]);
				}
			}

			adrLines = linesTemp;
			linesTemp = null;

			// PARSE
			var reArr:Array;

			// Name
			reArr = (adrLines[0] as String).match(/^Name: (.+)/);
			if (reArr != null) {
				name = trimSpaces(reArr[1] as String);
			}

			// Country
			country = trimSpaces(adrLines[adrLines.length-1] as String);

			return { name: name, country: country };

		}

		private function getCntServices(cnt:String):Array {

			var srvs:Array = null;

			if (srvsObj[cnt] != null)
				srvs = srvsObj[cnt];

			return srvs;

			/*
			switch (cnt) {
				case "United States":
					srvs.push("usps");
				break;
				case "Canada":
					srvs.push("ca");
				break;
				case "United Kingdom":
					srvs.push("gb_post");
					srvs.push("gb_post_det");
				break;
				case "Germany":
					srvs.push("dhl_ger_en");
				break;
				case "Spain":
					srvs.push("esp");
				break;
				case "Italy":
					srvs.push("it_post");
				break;
				case "Australia":
					srvs.push("aus");
				break;
				case "Slovakia":
					srvs.push("sk_post_en");
				break;
				case "Slovenia":
					srvs.push("si");
				break;
				case "Brazil":
					srvs.push("bra_en");
				break;
				case "Switzerland":
					srvs.push("swi");
				break;
				case "Chile":
					srvs.push("cl_correos");
				break;
				case "Czech Republic":
					srvs.push("cz_post_en");
				break;
				case "Denmark":
					srvs.push("dk");
				break;
				case "Finland":
					srvs.push("fi");
				break;
				case "France":
					srvs.push("fr_lap");
				break;
				case "Croatia":
					srvs.push("hr");
					srvs.push("hr_post");
				break;
				case "Hungary":
					srvs.push("hu");
				break;
				case "Ireland":
					srvs.push("ie");
					srvs.push("ie_post");
				break;
				case "Japan":
					srvs.push("jap");
				break;
				case "Korea":
					srvs.push("kor");
				break;
				case "Latvia":
					srvs.push("lv_en");
				break;
				case "Mexico":
					srvs.push("mx");
					srvs.push("mx_dhl");
				break;
				case "Netherlands":
					srvs.push("nl_post");
					srvs.push("nl_dhl");
				break;
				case "Norway":
					srvs.push("no");
				break;
				case "New Zealand":
					srvs.push("nz");
				break;
				case "Peru":
					srvs.push("pe_post");
				break;
				case "Poland":
					srvs.push("pl");
					srvs.push("pl_dhl");
				break;
				case "Portugal":
					srvs.push("pt_post");
				break;
				case "Sweden":
					srvs.push("se_dhl");
					srvs.push("se_post");
				break;
				case "Singapore":
					srvs.push("sg_post");
				break;
				case "Thailand":
					srvs.push("thai");
				break;
				case "Israel":
					srvs.push("isl");
				break;
				default:
					srvs = null;
				break;
			}

			return srvs;
			*/
		}

		private function parseServicesFile(fileString:String):Object {

			var file:String = fileString;
			var lineEnding:String = "\r\n"; // Windows style

			// > check CRLF. if not > error
			if (file.search(lineEnding) == -1) {
				outputLogLine("Ошибка в файле с сервисами: Wrong Line Ending", COLOR_BAD);
				return null;
			}

			var srvs:Object = {};
			var tmpArr:Array = []; // [!] Not in use
			var reArr:Array = [];
			var re1:RegExp = /^(.+) \[(.+)\]/;
			var re2:RegExp = /, ?/;

			var cnt:String;
			var srvsStr:String;
			var srvsArr:Array = [];

			var arr:Array = file.split(lineEnding);

			for each (var line:String in arr) {

				// Comment
				if (line.search(/^# ?/) != -1)
					continue;

				// Empty Line
				if (line == "")
					continue;

				if (line.search(re1) == -1) {
					outputLogLine("Ошибка в строке файла с сервисами: " + line, COLOR_BAD);
					continue;
				}

				reArr = line.match(re1);
				cnt = reArr[1];
				srvsStr = reArr[2];
				srvsArr = [];

				if (srvsStr.search(re2) == -1) {

					srvsArr.push(srvsStr);

				} else {

					srvsArr = srvsStr.split(re2);

				}

				srvs[cnt] = clone(srvsArr) as Array;
			}

			function clone(source:Object):* {
				var myBA:ByteArray = new ByteArray();
				myBA.writeObject(source);
				myBA.position = 0;
				return(myBA.readObject());
			}

			return srvs;

		}

		private function printDate(noTime:Boolean = false):String {
			var d:Date = new Date();
			var dtf:DateTimeFormatter = new DateTimeFormatter("ru-RU");
			var dstr:String;

			if (noTime) {
				dtf.setDateTimePattern("yyyy-MM-dd"); // 2017-03-20
				dstr = dtf.format(d);
				return dstr;
			}

			dtf.setDateTimePattern("yyyy-MM-dd @ HH:mm:ss"); // 2016-01-28T07:35:23
			dstr = dtf.format(d);
			dstr = dstr.replace(/\s@\s/, "T");
			return dstr;
		}

		private function trimSpaces(str:String):String {
			var ret:String = str.replace(/^\s*(.*?)\s*$/, "$1");
			return ret;
		}

		/**
		 * Paints an HTML-text to hex-color (Format: #000000) and returns HTML-formatted string
		 * @param color Hex-color of paint (Format: #000000)
		 * @param tx Text to be painted
		 * @return
		 */
		private function colorText(color:String, tx:String):String {
			return "<font color=\"" + color + "\">" + tx + "</font>";
		}

		private const srvsFileDefaultContent:String = "# ФОРМАТ:\r\n# Country [serv]\r\n# Country [serv1, serv2, ...]\r\n\r\nUnited States [usps]\r\nCanada [ca]\r\nUnited Kingdom [gb_post, gb_post_det]\r\nGermany [dhl_ger_en]\r\nSpain [esp]\r\nItaly [it_post]\r\nAustralia [aus]\r\nSlovakia [sk_post_en]\r\nSlovenia [si]\r\nBrazil [bra_en]\r\nSwitzerland [swi]\r\nChile [cl_correos]\r\nCzech Republic [cz_post_en]\r\nDenmark [dk]\r\nFinland [fi]\r\nFrance [fr_lap]\r\nCroatia [hr, hr_post]\r\nHungary [hu]\r\nIreland [ie, ie_post]\r\nJapan [jap]\r\nKorea [kor]\r\nLatvia [lv_en]\r\nMexico [mx, mx_dhl]\r\nNetherlands [nl_post, nl_dhl]\r\nNorway [no]\r\nNew Zealand [nz]\r\nPeru [pe_post]\r\nPoland [pl, pl_dhl]\r\nPortugal [pt_post]\r\nSweden [se_dhl, se_post]\r\nSingapore [sg_post]\r\nThailand [thai]\r\nIsrael [isl]";

	}

}