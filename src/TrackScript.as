package
{
	import com.childoftv.xlsxreader.Worksheet;
	import com.childoftv.xlsxreader.XLSXLoader;
	import fl.controls.Button;
	import fl.controls.TextInput;
	import flash.desktop.Clipboard;
	import flash.desktop.ClipboardFormats;
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
	import flash.net.FileReference;
	import flash.system.Capabilities;
	import flash.text.TextField;
	import flash.text.TextFormat;
	import flash.ui.Keyboard;
	import flash.utils.ByteArray;
	
	/**
	 * ...
	 * @author Tim Yusupov
	 */
	public class TrackScript extends Sprite
	{
		private var main:Main;
		private var ui:UI;
		private var win:NativeWindow;
		
		private const undoDirName:String = "undo";
		private const UNDO_NUMBER:int = 4;
		private var undoDir:File;
		
		private var btnFWMode:ModeButton;
		private var btnCantonMode1:ModeButton;
		private var btnCantonMode2:ModeButton;
		
		private const COLOR_BAD:String = "#CC171C"; // Red
		private const COLOR_SUCCESS:String = "#189510"; // Green
		private const COLOR_WARN:String = "#CB5815"; // Orange
		private const COLOR_SPECIAL:String = "#0075BF"; // Blue
		
		private const PRCMODE_FRONTWINNER:int = 1;
		private const PRCMODE_CANTON_WH_1:int = 2;
		private const PRCMODE_CANTON_WH_2:int = 3;
		
		private var prcMode:int;
		private var currentPrcModeButton:ModeButton;
		private var xlFilePathError:Boolean = false;
		private var xmlFilePathError:Boolean = false;
		private var devFlag:Boolean = false;
		private var runsCount:int;
		
		public function TrackScript():void
		{
			stage ? init() : addEventListener(Event.ADDED_TO_STAGE, init);
		}
		
		private function init(e:Event = null):void
		{
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
			
			win.addEventListener(NativeWindowBoundsEvent.MOVE, function(e:NativeWindowBoundsEvent):void
			{
				main.settings.setKey(Settings.winPos, win.x + ":" + win.y);
			});
			
			win.addEventListener(Event.CLOSING, function(e:Event):void
			{
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
			
			// Mode buttons
			btnFWMode = new ModeButton(ui.btnFWMode, PRCMODE_FRONTWINNER);
			btnFWMode.addEventListener("click", onModeButtonClick);
			btnCantonMode1 = new ModeButton(ui.btnCantonMode1, PRCMODE_CANTON_WH_1);
			btnCantonMode1.addEventListener("click", onModeButtonClick);
			btnCantonMode2 = new ModeButton(ui.btnCantonMode2, PRCMODE_CANTON_WH_2);
			btnCantonMode2.addEventListener("click", onModeButtonClick);
			
			function onModeButtonClick(e:Event):void 
			{
				switchMode((e.target as ModeButton).linkedMode);
			}
			
			// Init mode
			switchMode(main.settings.getKey(Settings.prcMode) as int);
			
			xlFile = new File();
			xmlFile = new File();
			xlFile.addEventListener(Event.SELECT, onFileSelect);
			xmlFile.addEventListener(Event.SELECT, onFileSelect);
			
			loaderInfo.uncaughtErrorEvents.addEventListener(UncaughtErrorEvent.UNCAUGHT_ERROR, onUncaughtError);
			
			// Check services file
			var srvsFile:File = File.applicationStorageDirectory.resolvePath("services.txt");
			if (!srvsFile.exists)
			{
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
			
			// Resolve Undo dir
			undoDir = File.applicationStorageDirectory.resolvePath(undoDirName);
			
			// Check dev marker-file
			if (File.applicationStorageDirectory.resolvePath("dev").exists)
				devFlag = true;
				
			// Reset number of runs of the script stat
			runsCount = 0;
		}
		
		private function onUncaughtError(e:UncaughtErrorEvent):void
		{
			if (e.error is Error)
			{
				var error:Error = e.error as Error;
				outputLogLine("Системная ошибка: " + error.message, COLOR_BAD);
				outputLogLine("Требуется проинформировать об этом создателя", COLOR_BAD);
			}
		}
		
		private function btnStartClick(e:MouseEvent):void
		{
			start();
		}
		
		private function keyDown(e:KeyboardEvent):void
		{
			// ESC
			if (e.keyCode == Keyboard.ESCAPE)
			{
				main.exitApp();
			}
			
			else
				
			// ENTER
			if (e.keyCode == Keyboard.ENTER)
			{
				start();
			}
		}
		
		private function onTfChange(e:Event):void
		{
			var tf:TextInput = e.target as TextInput;
			
			if (tf == ui.tfXlFile)
			{
				main.settings.setKey(Settings.sourceExcelFile, ui.tfXlFile.text);
				
				try
				{
					xlFile.nativePath = ui.tfXlFile.text;
					ui.tfXlFile.htmlText = colorText("#000000", ui.tfXlFile.text);
					xlFilePathError = false;
				}
				
				catch (e:Error)
				{
					ui.tfXlFile.htmlText = colorText(COLOR_BAD, ui.tfXlFile.text);
					xlFilePathError = true;
				}
				
				if (!validateFileInput(ui.tfXlFile.textField, xlFile, /.xlsx$/))
				{
					xlFilePathError = true;
				}
				else
				{
					xlFilePathError = false;
				}
			}
			
			else if (tf == ui.tfXmlFile)
			{
				main.settings.setKey(Settings.trackCheckerDataFile, ui.tfXmlFile.text);
				
				try
				{
					xmlFile.nativePath = ui.tfXmlFile.text;
					ui.tfXmlFile.htmlText = colorText("#000000", ui.tfXmlFile.text);
					xmlFilePathError = false;
				}
				
				catch (e:Error)
				{
					ui.tfXmlFile.htmlText = colorText(COLOR_BAD, ui.tfXmlFile.text);
					xmlFilePathError = true;
				}
				
				if (!validateFileInput(ui.tfXmlFile.textField, xmlFile, /.xml$/))
				{
					xmlFilePathError = true;
				}
				else
				{
					xmlFilePathError = false;
				}
			}
		}
		
		private function onFileSelect(e:Event):void
		{
			var file:File = e.target as File;
			
			if (file == xlFile)
			{
				ui.tfXlFile.text = file.nativePath;
				ui.tfXlFile.dispatchEvent(new Event("change"));
			}
			
			else if (file == xmlFile)
			{
				ui.tfXmlFile.text = file.nativePath;
				ui.tfXmlFile.dispatchEvent(new Event("change"));
			}
		}
		
		private function btnDialogClick(e:MouseEvent):void
		{
			var btn:Button = e.target as Button;
			
			if (btn == ui.btnXlDialog)
			{
				xlFile.browseForOpen("Файл Excel 2007 с информацией о треках в формате склада Кантона",
					[new FileFilter("Excel 2007", "*.xlsx")]);
			}
			
			else if (btn == ui.btnXmlDialog)
			{
				xmlFile.browseForOpen("Файл data.xml от TrackChecker",
					[new FileFilter("Файл XML", "*.xml")]);
			}
		}
		
		private function validateFileInput(inputTF:TextField, file:File, fileExtensionTpl:RegExp):Boolean
		{
			// Empty path
			if (inputTF.text == "")
				return false;
			
			// Not that extension
			if (inputTF.text.search(fileExtensionTpl) == -1)
			{
				inputTF.htmlText = colorText(COLOR_BAD, inputTF.text);
				return false;
			}
			else
			{
				inputTF.htmlText = colorText("#000000", inputTF.text);
			}
			
			// Not existing
			try
			{
				if (!file.exists)
				{
					inputTF.htmlText = colorText(COLOR_BAD, inputTF.text);
					return false;
				}
				else
				{
					inputTF.htmlText = colorText("#000000", inputTF.text);
				}
			}
			
			catch (e:Error)
			{
				trace("АШИБКА");
			}
			
			return true;
		}
		
		private function outputLogLine(tx:String, color:String = null):void
		{
			if (tx == "=")
				tx = "=============================";
			
			ui.taOutput.htmlText += (color != null ? colorText(color, tx) : tx) + "\n";
			ui.taOutput.verticalScrollPosition = ui.taOutput.maxVerticalScrollPosition;
		}
		
		private function switchMode(modeValue:int):void 
		{
			var modeBtn:ModeButton;
			
			switch (modeValue) 
			{
				case PRCMODE_FRONTWINNER:
					modeBtn = btnFWMode;
					break;
					
				case PRCMODE_CANTON_WH_1:
					modeBtn = btnCantonMode1;
					break;
					
				case PRCMODE_CANTON_WH_2:
					modeBtn = btnCantonMode2;
					break;
					
				default:
					throw new Error("Out of possible modes");
					return;
					break;
			}
			
			prcMode = modeValue;
			main.settings.setKey(Settings.prcMode, prcMode);
			if (currentPrcModeButton != null)
				currentPrcModeButton.active = false;
			currentPrcModeButton = modeBtn;
			currentPrcModeButton.active = true;
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
		
		// Common stats for all modes
		private var tracksCount:uint;
		private var existingTracksCount:uint;
		
		private var srvsObj:Object = {};
		
		private function start():void
		{
			if (xmlFilePathError)
			{
				outputLogLine("Ошибка в пути к файлу data.xml. Запуск невозможен.", COLOR_BAD);
				return;
			}
			
			ui.taOutput.text = "";
			
			// Increase number of runs stat
			runsCount++;
			
			outputLogLine("Запуск #" + runsCount + " [" + getFormattedDate("HH:mm:ss") + "]");
			
			// XML File Reading Start
			xmlFile.nativePath = ui.tfXmlFile.text;
			fst = new FileStream();
			
			fst.addEventListener(ProgressEvent.PROGRESS, onXMLFileLoadingProgress);
			fst.addEventListener(Event.COMPLETE, onXMLFileLoadingDone);
			fst.openAsync(xmlFile, FileMode.READ);
			outputLogLine("Загрузка файла data.xml от TrackChecker");
		}
		
		private function onXMLFileLoadingProgress(e:ProgressEvent):void
		{
			trace("Loading XML File. Progress:",
				e.bytesLoaded + " / " + e.bytesTotal, Math.floor(e.bytesLoaded / e.bytesTotal * 100) + "%");
		}
		
		private function onXMLFileLoadingDone(e:Event):void 
		{
			// Preparing data.xml
			trace("XML Data File: reading done");
			fst.removeEventListener(ProgressEvent.PROGRESS, onXMLFileLoadingProgress);
			fst.removeEventListener(Event.COMPLETE, onXMLFileLoadingDone);
			
			trace("Bytes Available " + fst.bytesAvailable);
			xmlString = fst.readUTFBytes(fst.bytesAvailable)
			fst.close();
			
			trace("XML String Length:", xmlString.length);
			
			dataXml = new XML(xmlString);
			
			var rootNodesCount:uint = dataXml.groups.children().length();
			trace("Nodes Count:", rootNodesCount);
			
			if (dataXml.@maxid == "" || dataXml.@maxid == null)
			{
				outputLogLine("Неверный формат \"data.xml\"", COLOR_BAD);
				return;
			}
			
			maxID = maxIDDefault = uint(dataXml.@maxid);
			trace("MAX ID:", dataXml.@maxid);
			
			outputLogLine("Имеющиеся группы в корне: " + String(rootNodesCount));
			
			// Capture current date to use it in XML output
			currentDate = printDate();
			
			// Calling the right function based on current state of mode
			switch (prcMode) 
			{
				// Frontwinner provider
				case PRCMODE_FRONTWINNER:
					processFrontwinnerMode();
					break;
					
				// Canton warehouse modes 1 and 2 (one processing function)
				case PRCMODE_CANTON_WH_1:
				case PRCMODE_CANTON_WH_2:
					startLoadingCantonExcelFile();
					break;
					
				default:
					throw new Error("Out of possible modes");
					break;
			}
		}
		
		private function startLoadingCantonExcelFile():void
		{
			// Excel file check
			if (xlFilePathError)
			{
				outputLogLine("Ошибка в пути к файлу Excel. Запуск невозможен.", COLOR_BAD);
				return;
			}
			
			// Load excel file (Canton warehouse format)
			xlLoader = new XLSXLoader();
			xlLoader.addEventListener(Event.COMPLETE, processCantonExcelModes1and2);
			xlLoader.load(ui.tfXlFile.text);
			
			outputLogLine(
				"Загрузка файла Excel в формате склада Кантона. Формат #" +
				(prcMode == PRCMODE_CANTON_WH_1 ? "1" : "2"), 
				COLOR_SPECIAL
			);
		}
		
		private function processCantonExcelModes1and2(e:Event):void
		{
			xlLoader.removeEventListener(Event.COMPLETE, processCantonExcelModes1and2);
			xlSheet = xlLoader.worksheet("Sheet1");
			
			if (xlSheet.getCellValue("A1").search(/Parcel List/i) == -1)
			{
				outputLogLine("Неверный формат Excel", COLOR_BAD);
				return;
			}
			
			// Local shortcuts for Canton modes
			var md1:Boolean = (prcMode == PRCMODE_CANTON_WH_1);
			var md2:Boolean = (prcMode == PRCMODE_CANTON_WH_2);
			
			/**
			 * Parsing Canton Excel Format
			 * ================================================================================
			 */
			var active:Boolean;
			var row:int;
			var groupMode:Boolean;
			var groups:XML;
			var currentGroup:XML;
			var emptyLinesCount:int;
			var wholeLineIsEmpty:Boolean;
			
			// Local stats
			var existingGroupsCount:uint;
			
			// Reset stats
			tracksCount = existingGroupsCount = existingTracksCount = 0;
			
			row = 3;
			groupMode = false;
			emptyLinesCount = 0;
			wholeLineIsEmpty = false;
			
			groups = <track-script-output/>;
			
			var groupColVal:String;
			var trackColVal:String;
			var adrColVal:String; // Mode 1 only
			var nameColVal:String; // Mode 2 only
			var cntColVal:String; // Mode 2 only
			
			var trackCol:String;
			var trackColumns:Array = md1 ? ["J"] : ["H", "P", "Q"];
			
			var groupName:String;
			var name:String;
			var track:String;
			var country:String;
			
			// Find column with tracks
			var colVal:String;
			for each (var trco:String in trackColumns)
			{
				colVal = xlSheet.getCellValue(trco + row);
				colVal = trimSpaces(colVal);
				if (colVal.search(/^[a-zA-Z]{2}\d{9}(CN|cn|Cn|cN)?$/) != -1)
				{
					trackCol = trco;
					break;
				}
			}
			
			if (trackCol == null)
			{
				outputLogLine("Ошибка: не найдена колонка с треками", COLOR_BAD);
				return;
			}
			
			outputLogLine("Колонка с треками: " + trackCol);
			
			// Date parse (Mode 2 only)
			if (md2)
			{
				var tableDate:String = trimSpaces(xlSheet.getCellValue("A3"));
				var tableDateRegExPattern:RegExp = /(\d{2})\.(\d{2})\.(\d{4})/;
				
				if (tableDate.search(tableDateRegExPattern) == -1)
				{
					tableDate = null;
					outputLogLine("Не найдена дата в файле. Текущая дата записана в событие «Added» всех треков", COLOR_WARN);
				}
				
				else
				{
					outputLogLine("Дата " + tableDate + " из файла записана в событие «Added» всех треков", COLOR_SPECIAL);
					var re:Array = tableDate.match(tableDateRegExPattern);
					tableDate = re[3] + "-" + re[2] + "-" + re[1];
				}
			}
			
			active = true;
			
			while (active)
			{
				// Reset some variables
				wholeLineIsEmpty = false;
				
				groupColVal = xlSheet.getCellValue("C" + row);
				trackColVal = xlSheet.getCellValue(trackCol + row);
				
				if (md1)
				{
					adrColVal = xlSheet.getCellValue("H" + row);
				}
				
				else if (md2)
				{
					nameColVal = xlSheet.getCellValue("K" + row);
					cntColVal = xlSheet.getCellValue("J" + row);
				}
				
				// ================================================================================
				
				if (md1)
				{
					if (groupColVal == "" && trackColVal == "" && adrColVal == "")
					{
						wholeLineIsEmpty = true;
					}
					
					if (adrColVal == "" || adrColVal == null)
					{
						!wholeLineIsEmpty && outputLogLine("Пустой адрес на строке " + row, COLOR_WARN);
					}
					else
					{
						var adrObj:Object = parseAddressCell(adrColVal);
						name = adrObj.name;
						country = adrObj.country;
					}
				}
				
				else if (md2)
				{
					if (groupColVal == "" && trackColVal == "" && cntColVal == "")
					{
						wholeLineIsEmpty = true;
					}
					
					name = trimSpaces(nameColVal);
					country = trimSpaces(cntColVal);
				}
				
				// ================================================================================
				
				if (trackColVal == "")
				{
					!wholeLineIsEmpty && outputLogLine("Пустой трек на строке " + row, COLOR_WARN);
				}
				
				if (trackColVal != "" && emptyLinesCount == 1)
				{
					emptyLinesCount = 0;
				}
				
				row++;
				
				if (emptyLinesCount > 1)
				{
					active = false;
					break;
				}
				
				if (trackColVal == "")
				{
					emptyLinesCount++;
					continue;
				}
				
				if (groupColVal != "")
				{
					groupMode = true;
					groupName = trimSpaces(groupColVal);
					
					currentGroup = createXmlGroup(groupName);
					groups.appendChild(currentGroup);
				}
				
				track = trimSpaces(trackColVal);
				//name = trimSpaces(nameColVal);
				//country = trimSpaces(cntColVal);
				
				// [!] Special date for Mode 2
				var xmlTrack:XML = createXmlTrack(name, track, country, (md2 && tableDate != null) ? tableDate : null);
				currentGroup.appendChild(xmlTrack);
				tracksCount++;
			}
			
			// Write result of script work to output.xml
			groups.@date = getFormattedDate("dd.MM.yyyy HH:mm:ss");
			groups.@version = main.version;
			groups.@format = md1 ? "Canton #1" : "Canton #2";
			groups.@excel = ui.tfXlFile.text;
			
			writeToOutputFile(groups);
			
			/**
			 * Filling source XML (dataXml) with generated groups
			 * ================================================================================
			 */
			// Iterate through our groups
			for each (var grp:XML in groups.groups)
			{
				trace("Group:", grp.@desc);
				
				// Check if this group already exists
				// If exists
				var x:* = dataXml.groups.(@id == 0)..groups.(@desc == grp.@desc);
				if (x.length() > 0)
				{
					// Compare tracks in existing group to our current group
					for each (var tr:XML in groups.groups.(@desc == grp.@desc).track)
					{
						trace("Track:", tr.@track);
						
						// Track Existence Check
						// If found duplicate > skip this track
						x = dataXml.groups.(@id == 0)..groups.(@desc == grp.@desc).track.(@desc == tr.@desc && @track == tr.@track);
						if (x.length() > 0)
						{
							main.logRed("Track Duplicate Found: " + tr.@desc + " " + tr.@track);
							existingTracksCount++;
							continue;
						}
						
						// If not found > add this track to existing group
						else
						{
							for each (var subGroup:XML in dataXml.groups.(@id == 0)..groups.(@desc == grp.@desc))
							{
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
			if (existingTracksCount > 0) outputLogLine("Найденные дубли треков: " + existingTracksCount, COLOR_WARN);
			
			if (tracksCount == existingTracksCount && tracksCount > 0)
			{
				outputLogLine("Похоже кто-то прогнал меня два раза", COLOR_BAD);
			}
			
			if (Capabilities.isDebugger || devFlag)
			{
				outputLogLine("Готово", COLOR_SUCCESS);
				return;
			}
			
			/**
			 * Write Back to XML File
			 * ================================================================================
			 */
			writeBackToXMLFile();
			
			outputLogLine("Готово", COLOR_SUCCESS);
		}
		
		private function processFrontwinnerMode():void 
		{
			/*
			Главный алгоритм
			> Parse text with information from clipboard
			> Prepare XML-tree based on parsed records to be inserted in main XML
			> Add script's result to in-memory data.xml
			> Write script's result to output.xml
			> Show some stats
			> Write back in-memory data to data.xml file
			*/
			
			if (!Clipboard.generalClipboard.hasFormat(ClipboardFormats.TEXT_FORMAT))
			{
				outputLogLine("В буфере обмена не найден текст", COLOR_BAD);
				return;
			}
			
			outputLogLine("Чтение текста в формате Frontwinner из буфера обмена", COLOR_SPECIAL);
			var tx:String = Clipboard.generalClipboard.getData(ClipboardFormats.TEXT_FORMAT) as String;
			
			/**
			 * Split Text to Array of Lines
			 * ================================================================================
			 */
			var ctrlCharPattern:RegExp = /(\r|\n|\r\n)/;
			
			// Check: empty or one line
			if (tx.length < 1 || tx.search(ctrlCharPattern) == -1) {
				outputLogLine("Ошибка: одна строчка в буфере обмена", COLOR_BAD);
				return;
			}
			
			var txAr:Array = [];
			var linesTemp:Array = [];

			// Разделить по строкам
			txAr = tx.split(ctrlCharPattern);

			var i:int;

			// Отчистить от управляющих символов
			for (i = 0; i < txAr.length; i++) {
				if ((txAr[i] as String).search(ctrlCharPattern) == -1) {
					linesTemp.push(txAr[i]);
				}
			}

			txAr = linesTemp;
			linesTemp = [];

			// Отчистить от пустых символов
			for (i = 0; i < txAr.length; i++) {
				if ((txAr[i] as String).length != 0 || (txAr[i] as String) != "") {
					linesTemp.push(txAr[i]);
				}
			}

			txAr = linesTemp;
			linesTemp = null;
									
			/**
			 * Parse Provider Text (Frontwinner)
			 * ================================================================================
			 */
			var reAr:Array = []; // Temp array for RegEx operations
			var currentTrackRecord:Object;
			var tmpRecordSourceLines:Vector.<String>;
			var allTrackRecords:Vector.<Object>;
			
			var totalRecordsCount:int = 0;
			var trackRecordsCount:int = 0;
			var notTrackRecordsCount:int = 0;
			var invalidTrackRecordsCount:int = 0;
			
			var active:Boolean = false;
			var idx:int = 0; // Current index in text array
			var l:String; // Current processed line in cycle
			var linesInRecord:int = 0;
			var isTrackRecord:Boolean;
			var isPrevRecordTrack:Boolean;
			
			var headerLineMark:RegExp = /^([\d-]+) ?#/;
			var dateInHeaderTpl:RegExp = /^([\d-]+)(?= ?#)/;
			var trackRecordHeaderMark:RegExp = /Shipped$/i;
			
			// Frontwinner format match check
			if ((txAr[0] as String).search(headerLineMark) == -1 || txAr.length < 6)
			{
				outputLogLine("Ошибка: текст в буфере не походит на формат Frontwinner", COLOR_BAD);
				return;
			}
			
			// Start parsing
			active = true;
			allTrackRecords = new Vector.<Object>();
			
			// Reset common stats
			tracksCount = existingTracksCount = 0;
			
			// Go through every line of text step by step
			while (active)
			{
				// No more lines to handle
				if (idx == txAr.length)
				{
					// Finish final track record (if it was last)
					if (isTrackRecord && tmpRecordSourceLines.length != 0) 
						finishTrackRecord();
					
					// End of the parsing
					active = false;
					break;
				}
				
				l = txAr[idx] as String;
				idx++;
				
				// Header (header line mark occurrence)
				if (l.search(headerLineMark) != -1) 
				{
					isPrevRecordTrack = isTrackRecord;
					isTrackRecord = l.search(trackRecordHeaderMark) != -1 ? true : false;
					totalRecordsCount++;
					
					if (!isTrackRecord)
						notTrackRecordsCount++;
					
					if (totalRecordsCount > 1 && isPrevRecordTrack)
						finishTrackRecord();
					
					if (isTrackRecord)
					{
						initNewTrackRecord();
					}
					else
					{
						continue;
					}
				}
				
				// Not Header (ordinary line)
				else
				{
					// Check: header line should be first in text
					// otherwise first ordinary line without header before is skipped
					if (totalRecordsCount == 0)
						continue;
					
					// Check whether script is aware of filling track record
					if (isTrackRecord)
					{
						// Add the line to temp array of record lines
						// and continue to next
						tmpRecordSourceLines.push(l);
						continue;
					}
					else
					{
						continue; // Just for sake of code clarity
					}
				}
			}
			
			function initNewTrackRecord():void 
			{
				currentTrackRecord = {};
				tmpRecordSourceLines = new Vector.<String>();
				
				// Retrieve date
				reAr = l.match(dateInHeaderTpl); // From Header line (it's currently being processed)
				currentTrackRecord.date = reAr[0];
				
				trackRecordsCount++;
			}
			
			function finishTrackRecord():int 
			{
				linesInRecord = tmpRecordSourceLines.length;
				if (linesInRecord < 5) 
				{
					outputLogLine("Неверный формат блока " + trackRecordsCount, COLOR_WARN);
					currentTrackRecord = null; // Invalid record isn't added to final list of records (allRecords)
					invalidTrackRecordsCount++;
					return 1;
				}
				
				currentTrackRecord.track = trimSpaces(tmpRecordSourceLines[1]);
				currentTrackRecord.name = trimSpaces(tmpRecordSourceLines[2]);
				currentTrackRecord.country = trimSpaces(tmpRecordSourceLines[tmpRecordSourceLines.length-1]); // Last line
				allTrackRecords.push(currentTrackRecord);
				tracksCount++; // Track is considered handled only if record is valid
				return 0;
			}
			
			// Check
			if (allTrackRecords.length == 0) 
			{
				outputLogLine("Ошибка: ни одного блока не найдено", COLOR_BAD);
				return;
			}
			
			/**
			 * Form tracks in XML
			 * ================================================================================
			 */
			/*
			Алгоритм
			> Create XML-tree
			> Check 'Frontwinner' group existence in data.xml
			> 	* if found > add all tracks there
			>	* if not found > create such group; add all tracks there
			*/
			
			var newFrontwinnerGrp:XML = createXmlGroup("Frontwinner");
			var rec:Object
			for each (rec in allTrackRecords) 
			{
				newFrontwinnerGrp.appendChild(createXmlTrack(rec.name, rec.track, rec.country, rec.date));
			}
			
			var rootGroup:XMLList = dataXml.groups.(@id == 0);
			
			var x:*;
			var frontwinnerGrp:* = rootGroup..groups.(@desc == "Frontwinner");
			if (frontwinnerGrp.length() > 0)
			{
				var newTracks:XMLList = newFrontwinnerGrp.track;
				for each (var t:XML in newTracks)
				{
					// Track Existence Check
					// If found duplicate > skip this track
					x = frontwinnerGrp.track.(@desc == t.@desc && @track == t.@track);
					if (x.length() > 0)
					{
						main.logRed("Track Duplicate Found: " + t.@desc + " " + t.@track);
						existingTracksCount++;
						continue;
					}
					
					frontwinnerGrp.appendChild(t);
				}
			}
			else
			{
				rootGroup.prependChild(newFrontwinnerGrp);
			}
			
			/**
			 * Write to Output File
			 * ================================================================================
			 */
			var output:XML = <track-script-output/>;
			output.appendChild(newFrontwinnerGrp);
			output.@date = getFormattedDate("dd.MM.yyyy HH:mm:ss");
			output.@version = main.version;
			output.@format = "Frontwinner";
			writeToOutputFile(output);
			
			/**
			 * Show Stats
			 * ================================================================================
			 */
			outputLogLine("Обработано треков: " + tracksCount + ", всего найдено блоков: " + totalRecordsCount);
			if (invalidTrackRecordsCount > 0) outputLogLine("Количество неверных блоков: " + invalidTrackRecordsCount, COLOR_BAD);
			if (notTrackRecordsCount > 0) outputLogLine("Пропущенные блоки: " + notTrackRecordsCount, COLOR_WARN);
			if (existingTracksCount > 0) outputLogLine("Найденные дубли треков: " + existingTracksCount, COLOR_WARN);
			
			if (tracksCount == existingTracksCount && tracksCount > 0)
			{
				outputLogLine("Похоже кто-то прогнал меня два раза", COLOR_BAD);
			}
			
			if (Capabilities.isDebugger || devFlag)
			{
				outputLogLine("Готово", COLOR_SUCCESS);
				return;
			}
			
			/**
			 * Write Back to XML File
			 * ================================================================================
			 */
			writeBackToXMLFile();
			
			outputLogLine("Готово", COLOR_SUCCESS);
		}
		
		private function writeBackToXMLFile():void 
		{
			// Backup untouched dataXml file before update it
			var undoFileBackup:FileReference;
			var date:Date = new Date();
			var dtf:DateTimeFormatter = new DateTimeFormatter("ru-RU");
			var dstr:String;
			
			dtf.setDateTimePattern("dd.MM.yy-HH.mm.ss");
			dstr = dtf.format(date);
			undoFileBackup = undoDir.resolvePath("launch-" + dstr + ".xml");
			
			xmlFile.copyToAsync(undoFileBackup, true);
			xmlFile.addEventListener(Event.COMPLETE, writeBackToXMLFile_p2);
		}
		
		private function writeBackToXMLFile_p2(e:Event):void 
		{
			xmlFile.removeEventListener(Event.COMPLETE, writeBackToXMLFile_p2);
			
			dataXml.@maxid = maxID; // Update MaxID
			XML.prettyPrinting = false;
			var outputStr:String = dataXml.toXMLString();
			
			fst.openAsync(xmlFile, FileMode.WRITE);
			fst.writeUTFBytes(outputStr);
			fst.close();
			
			outputLogLine("Добавлено в TrackChecker", COLOR_SUCCESS);
			checkUndoFilesForCleanup();
		}
				
		private function checkUndoFilesForCleanup():void 
		{
			main.logRed("Checking undo files for cleanup");
			var dirContents:Array = undoDir.getDirectoryListing();
			
			if (dirContents.length == 0 || dirContents.length <= UNDO_NUMBER)
				return;
			
			var sortedFiles:Array = [];
			
			for each (var f:File in dirContents)
			{
				sortedFiles.push({ts: f.creationDate.time, file: f});
			}
			
			sortedFiles.sortOn("ts", Array.NUMERIC | Array.DESCENDING);
			
			if (sortedFiles.length >= UNDO_NUMBER)
			{
				var removed:Object;
				var dif:Number = sortedFiles.length - UNDO_NUMBER;
				for (var i:int = 0; i < dif; i++) 
				{	
					removed = sortedFiles.pop();
					removeFile(removed.file as File);
				}
			}
			
			function removeFile(f:File):void
			{
				f.deleteFileAsync();
				main.logRed("File \"" + f.name + "\" has been removed");
			}
		}
		
		private function writeToOutputFile(xml:XML):void 
		{
			XML.prettyPrinting = true;
			XML.prettyIndent = 4;
			var outputFile:File = File.applicationStorageDirectory.resolvePath("output.xml");
			fst.openAsync(outputFile, FileMode.WRITE);
			fst.writeUTFBytes(xml.toXMLString());
			fst.close();
		}
		
		private function createXmlTrack(name:String, track:String, country:String, addedEventSpecialDate:String = null):XML
		{
			var xmlTrack:XML = <track/>;
			var xmlTrackServs:XML = <servs/>;
			var xmlServ:XML;
			
			xmlTrack.@id = ++maxID; 
			xmlTrack.@desc = name;
			xmlTrack.@crdt = currentDate;
			xmlTrack.@track = track;
			
			xmlTrackServs.@id = ++maxID;
			xmlTrackServs.@crdt = currentDate;
			
			var servAliases:Array = ["china", "china_alt", "china_ems"];
			var servAlias:String;
			
			for each (servAlias in servAliases)
			{
				xmlServ = <serv/>;
				xmlServ.@id = ++maxID;
				xmlServ.@crdt = currentDate;
				xmlServ.@serv = servAlias;
				xmlServ.@selected = 1;
				xmlTrackServs.appendChild(xmlServ);
			}
			
			var specialServs:Array = getCntServices(country);
			
			if (specialServs != null)
			{
				for each (servAlias in specialServs)
				{
					xmlServ = <serv/>;
					xmlServ.@id = ++maxID;
					xmlServ.@crdt = currentDate;
					xmlServ.@serv = servAlias;
					xmlServ.@selected = 1;
					xmlTrackServs.appendChild(xmlServ);
				}
			}
			
			// Special event "Added". Appended to every track
			var xmlEventAdded:XML = <event/>;
			xmlEventAdded.@id = ++maxID;
			xmlEventAdded.@crdt = addedEventSpecialDate == null ? currentDate : addedEventSpecialDate + "T00:00:00";
			xmlEventAdded.@desc = "Added";
			xmlEventAdded.@udt = addedEventSpecialDate == null ? printDate(true) : addedEventSpecialDate;
			
			xmlTrack.appendChild(xmlTrackServs);
			xmlTrack.appendChild(xmlEventAdded);
			
			return xmlTrack;
		}
		
		private function createXmlGroup(grpName:String):XML
		{
			var g:XML = <groups/>;
			g.@id = ++maxID;
			g.@desc = grpName;
			g.@crdt = currentDate;
			return g;
		}
		
		private function parseAddressCell(adrCellVal:String):Object
		{
			var ctrlCharPattern:RegExp = /(\r|\n|\r\n)/;
			var name:String;
			var country:String;
			
			var adrColVal:String = trimSpaces(adrCellVal);
			
			// Check: empty or one line
			if (adrColVal.length < 1 || adrColVal.search(ctrlCharPattern) == -1)
			{
				// return
			}
			
			var adrLines:Array;
			var linesTemp:Array = [];
			
			// Разделить по строкам
			adrLines = adrColVal.split(ctrlCharPattern);
			
			var i:int;
			
			// Отчистить от управляющих символов
			for (i = 0; i < adrLines.length; i++)
			{
				if ((adrLines[i] as String).search(ctrlCharPattern) == -1)
				{
					linesTemp.push(adrLines[i]);
				}
			}
			
			adrLines = linesTemp;
			linesTemp = [];
			
			// Отчистить от пустых символов
			for (i = 0; i < adrLines.length; i++)
			{
				if ((adrLines[i] as String).length != 0 || (adrLines[i] as String) != "")
				{
					linesTemp.push(adrLines[i]);
				}
			}
			
			adrLines = linesTemp;
			linesTemp = null;
			
			// PARSE
			var reArr:Array;
			
			// Name
			reArr = (adrLines[0] as String).match(/^Name: (.+)/);
			if (reArr != null)
			{
				name = trimSpaces(reArr[1] as String);
			}
			
			// Country
			country = trimSpaces(adrLines[adrLines.length - 1] as String);
			
			return {name: name, country: country};
		}
		
		private function getCntServices(cnt:String):Array
		{
			var srvs:Array = null;
			
			if (srvsObj[cnt] != null)
				srvs = srvsObj[cnt];
			
			return srvs;
		}
		
		private function parseServicesFile(fileString:String):Object
		{
			var file:String = fileString;
			var lineEnding:String = "\r\n"; // Windows style
			
			// > check CRLF. if not > error
			if (file.search(lineEnding) == -1)
			{
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
			
			for each (var line:String in arr)
			{
				// Comment
				if (line.search(/^# ?/) != -1)
					continue;
				
				// Empty Line
				if (line == "")
					continue;
				
				if (line.search(re1) == -1)
				{
					outputLogLine("Ошибка в строке файла с сервисами: " + line, COLOR_BAD);
					continue;
				}
				
				reArr = line.match(re1);
				cnt = reArr[1];
				srvsStr = reArr[2];
				srvsArr = [];
				
				if (srvsStr.search(re2) == -1)
				{
					srvsArr.push(srvsStr);
				}
				else
				{
					srvsArr = srvsStr.split(re2);
				}
				
				srvs[cnt] = clone(srvsArr) as Array;
			}
			
			function clone(source:Object):*
			{
				var myBA:ByteArray = new ByteArray();
				myBA.writeObject(source);
				myBA.position = 0;
				return (myBA.readObject());
			}
			
			return srvs;
		}
		
		private function printDate(noTime:Boolean = false):String
		{
			var d:Date = new Date();
			var dtf:DateTimeFormatter = new DateTimeFormatter("ru-RU");
			var dstr:String;
			
			if (noTime)
			{
				dtf.setDateTimePattern("yyyy-MM-dd"); // 2017-03-20
				dstr = dtf.format(d);
				return dstr;
			}
			
			dtf.setDateTimePattern("yyyy-MM-dd @ HH:mm:ss"); // 2016-01-28T07:35:23
			dstr = dtf.format(d);
			dstr = dstr.replace(/\s@\s/, "T");
			return dstr;
		}
		
		private function trimSpaces(str:String):String
		{
			var ret:String = str.replace(/^\s*(.*?)\s*$/, "$1");
			return ret;
		}
		
		private function getFormattedDate(formatStr:String):String 
		{
			var d:Date = new Date();
			var dtf:DateTimeFormatter = new DateTimeFormatter("ru-RU");
			dtf.setDateTimePattern(formatStr);
			return dtf.format(d);
		}
		
		/**
		 * Paints an HTML-text to hex-color (Format: #000000) and returns HTML-formatted string
		 * @param color Hex-color of paint (Format: #000000)
		 * @param tx Text to be painted
		 * @return
		 */
		private function colorText(color:String, tx:String):String
		{
			return "<font color=\"" + color + "\">" + tx + "</font>";
		}
		
		private const srvsFileDefaultContent:String = "# ФОРМАТ:\r\n# Country [serv]\r\n# Country [serv1, serv2, ...]\r\n\r\nUnited States [usps]\r\nCanada [ca]\r\nUnited Kingdom [gb_post, gb_post_det]\r\nGermany [dhl_ger_en]\r\nSpain [esp]\r\nItaly [it_post]\r\nAustralia [aus]\r\nSlovakia [sk_post_en]\r\nSlovenia [si]\r\nBrazil [bra_en]\r\nSwitzerland [swi]\r\nChile [cl_correos]\r\nCzech Republic [cz_post_en]\r\nDenmark [dk]\r\nFinland [fi]\r\nFrance [fr_lap]\r\nCroatia [hr, hr_post]\r\nHungary [hu]\r\nIreland [ie, ie_post]\r\nJapan [jap]\r\nKorea [kor]\r\nLatvia [lv_en]\r\nMexico [mx, mx_dhl]\r\nNetherlands [nl_post, nl_dhl]\r\nNorway [no]\r\nNew Zealand [nz]\r\nPeru [pe_post]\r\nPoland [pl, pl_dhl]\r\nPortugal [pt_post]\r\nSweden [se_dhl, se_post]\r\nSingapore [sg_post]\r\nThailand [thai]\r\nIsrael [isl]";
	}
}