package
{
	import com.childoftv.xlsxreader.Worksheet;
	import com.childoftv.xlsxreader.XLSXLoader;
	import entities.ShenzhenPackage;
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
	import flash.utils.Dictionary;
	
	/**
	 * ...
	 * @author Tim Yusupov
	 */
	public class TrackScript extends Sprite
	{
		private var main:Main;
		private var ui:UI;
		private var win:NativeWindow;
		
		private const WEIGHT_STAT_MAX_RECORD_COUNT:int = 10;
		
		private const undoDirName:String = "undo";
		private const UNDO_NUMBER:int = 4;
		private var undoDir:File;
		
		private var btnFWMode:ModeButton;
		//private var btnCantonMode1:ModeButton;
		private var btnCantonMode2:ModeButton;
		private var btnShenzhen:ModeButton;
		
		// Semantic colors (HTML)
		private const COLOR_BAD:String = "#CC171C"; // Red
		private const COLOR_SUCCESS:String = "#189510"; // Green
		private const COLOR_WARN:String = "#CB5815"; // Orange
		private const COLOR_SPECIAL:String = "#0075BF"; // Blue
		
		// Colors for text on light background (HTML)
		public static const COLOR_TXLB_LIGHT_GREY:String = "#8B9398";
		
		private const PRCMODE_FRONTWINNER:int = 1;
		private const PRCMODE_CANTON_WH_1:int = 2; // Not in use
		private const PRCMODE_CANTON_WH_2:int = 3;
		private const PRCMODE_SHENZHEN:int = 4;
		
		private const postCodeTemplateForNetherlands:RegExp = /^([A-Za-z\d]{1,4}|\d{4,8})[-| ]*([A-Za-z\d]{1,4}|\d{4,8})/;
		
		private var prcMode:int;
		private var currentPrcModeButton:ModeButton;
		private var xlFilePathError:Boolean = false;
		private var xlDirPathError:Boolean = false;
		private var xmlFilePathError:Boolean = false;
		private var devFlag:Boolean = false;
		private var runsCount:int;
		
		private var fst:FileStream;
		private var xlLoader:XLSXLoader;
		private var xlSheet:Worksheet;
		private var xlColLetters:Array;
		private var xlFile:File;
		private var xlDir:File;
		private var xmlFile:File;
		private var dataXml:XML;
		private var xmlString:String;
		private var weightStatXml:XML;
		private var weightStatFile:File;
		
		private var maxIDDefault:uint;
		private var maxID:uint;
		private var currentDate:String;
		private var userDefinedDate:Date;
		
		// Common stats for all modes
		private var tracksCount:uint;
		private var existingTracksCount:uint;
		
		private var srvsObj:Object = {};
		
		public function TrackScript():void
		{
			stage ? init() : addEventListener(Event.ADDED_TO_STAGE, init);
		}
		
		private function init(e:Event = null):void
		{
			removeEventListener(Event.ADDED_TO_STAGE, init);
			loaderInfo.uncaughtErrorEvents.addEventListener(UncaughtErrorEvent.UNCAUGHT_ERROR, onUncaughtError);
			
			// Entry Point
			main = Main.ins;
			ui = new UI();
			addChild(ui);
			
			win = stage.nativeWindow;
			win.title = "TrackScript";
			ui.tfVer.text = "v" + main.version;
			
			var winPosStr:String = main.settings.getKey(Settings.winPos);
			var reResult:Array = winPosStr.match(/(-?\d+):(-?\d+)/);
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
			ui.tfXlDir.setStyle("textFormat", tfTextFormat);
			ui.tfXlDir.setStyle("disabledTextFormat", tfTextFormat);
			ui.tfXmlFile.setStyle("textFormat", tfTextFormat);
			ui.tfXmlFile.setStyle("disabledTextFormat", tfTextFormat);
			ui.taOutput.setStyle("textFormat", defTextFormat);
			ui.taOutput.setStyle("disabledTextFormat", defTextFormat);
			ui.btnStart.setStyle("textFormat", btnTextFormat);
			ui.btnStart.setStyle("disabledTextFormat", btnTextFormat);
			
			ui.nsDateDay.setStyle("textFormat", tfTextFormat);
			ui.nsDateDay.setStyle("disabledTextFormat", tfTextFormat);
			ui.nsDateDay.textField.setStyle("textFormat", tfTextFormat);
			ui.nsDateDay.textField.setStyle("disabledTextFormat", tfTextFormat);
			ui.nsDateMonth.setStyle("textFormat", tfTextFormat);
			ui.nsDateMonth.setStyle("disabledTextFormat", tfTextFormat);
			ui.nsDateMonth.textField.setStyle("textFormat", tfTextFormat);
			ui.nsDateMonth.textField.setStyle("disabledTextFormat", tfTextFormat);
			ui.nsDateYear.setStyle("textFormat", tfTextFormat);
			ui.nsDateYear.setStyle("disabledTextFormat", tfTextFormat);
			ui.nsDateYear.textField.setStyle("textFormat", tfTextFormat);
			ui.nsDateYear.textField.setStyle("disabledTextFormat", tfTextFormat);
			
			ui.tfXlFile.text = main.settings.getKey(Settings.sourceExcelFile);
			ui.tfXlDir.text = main.settings.getKey(Settings.excelDir);
			ui.tfXmlFile.text = main.settings.getKey(Settings.trackCheckerDataFile);
			ui.tfXlFile.addEventListener("change", onTfChange);
			ui.tfXlDir.addEventListener("change", onTfChange);
			ui.tfXmlFile.addEventListener("change", onTfChange);
			
			ui.taOutput.editable = false;
			
			ui.btnXlDialog.addEventListener(MouseEvent.CLICK, btnDialogClick);
			ui.btnXmlDialog.addEventListener(MouseEvent.CLICK, btnDialogClick);
			ui.btnXlDirDialog.addEventListener(MouseEvent.CLICK, btnDialogClick);
			ui.btnStart.addEventListener(MouseEvent.CLICK, btnStartClick);
			ui.btnStart.label = "З А П У С К";
			
			// Mode buttons
			btnFWMode = new ModeButton(ui.btnFWMode, PRCMODE_FRONTWINNER);
			btnFWMode.addEventListener("click", onModeButtonClick);
			//btnCantonMode1 = new ModeButton(ui.btnCantonMode1, PRCMODE_CANTON_WH_1);
			//btnCantonMode1.addEventListener("click", onModeButtonClick);
			btnCantonMode2 = new ModeButton(ui.btnCantonMode2, PRCMODE_CANTON_WH_2);
			btnCantonMode2.addEventListener("click", onModeButtonClick);
			btnShenzhen = new ModeButton(ui.btnShzMode, PRCMODE_SHENZHEN);
			btnShenzhen.addEventListener("click", onModeButtonClick);
			
			function onModeButtonClick(e:Event):void 
			{
				switchMode((e.target as ModeButton).linkedMode);
			}
			
			// Init mode
			switchMode(main.settings.getKey(Settings.prcMode) as int);
			
			xlFile = new File(ui.tfXlFile.text != "" ? ui.tfXlFile.text : null);
			xlDir = new File(ui.tfXlDir.text != "" ? ui.tfXlDir.text : null);
			xmlFile = new File(ui.tfXmlFile.text != "" ? ui.tfXmlFile.text : null);
			xlFile.addEventListener(Event.SELECT, onFileSelect);
			xlDir.addEventListener(Event.SELECT, onFileSelect);
			xmlFile.addEventListener(Event.SELECT, onFileSelect);
			
			xlColLetters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ"]; 
			
			// Date input init
			var d:Date = new Date(); // Current date
			ui.nsDateDay.textField.text = d.getDate().toString();
			ui.nsDateMonth.textField.text = String(d.getMonth() + 1);
			ui.nsDateYear.textField.text = String(d.getFullYear());
			
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
			
			// Weight statistics file
			checkWeightStatFile();
			
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
			
			else if (tf == ui.tfXlDir) 
			{
				xlDir.nativePath = ui.tfXlDir.text;
				
				if (!xlDir.exists || !xlDir.isDirectory) 
				{
					ui.tfXlDir.htmlText = colorText(COLOR_BAD, ui.tfXlDir.text);
					xlDirPathError = true;
					return;
				}
				else 
				{
					ui.tfXlDir.htmlText = colorText("#000000", ui.tfXlDir.text);
					xlDirPathError = false;
					main.settings.setKey(Settings.excelDir, ui.tfXlDir.text);
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
			
			else if (file == xlDir) 
			{
				ui.tfXlDir.text = file.nativePath;
				ui.tfXlDir.dispatchEvent(new Event("change"));
			}
		}
		
		private function btnDialogClick(e:MouseEvent):void
		{
			var btn:Button = e.target as Button;
			
			if (btn == ui.btnXlDialog)
			{
				xlFile.browseForOpen("Файл Excel 2007 с информацией о треках",
					[new FileFilter("Excel 2007", "*.xlsx")]);
			}
			
			else if (btn == ui.btnXmlDialog)
			{
				xmlFile.browseForOpen("Файл с данными TrackChecker",
					[new FileFilter("Файл XML", "*.xml")]);
			}
			
			else if (btn == ui.btnXlDirDialog) 
			{
				xlDir.browseForDirectory("Укажите папку с файлами *.xlsx");
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
		
		private function outputLogLine(tx:String, color:String = null, multiColorLine:Boolean = false):void
		{
			if (tx == "=")
				tx = "=============================";
			
			if (multiColorLine)	
				ui.taOutput.htmlText += tx + "\n";
			else
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
				
				/*	
				case PRCMODE_CANTON_WH_1:
					modeBtn = btnCantonMode1;
					break;
				*/	
					
				case PRCMODE_CANTON_WH_2:
					modeBtn = btnCantonMode2;
					break;
				
				case PRCMODE_SHENZHEN:
					modeBtn = btnShenzhen;
					break;
					
				default:
					main.logRed("Warning! Out of possible modes");
					modeValue = PRCMODE_FRONTWINNER; // Set default
					modeBtn = btnFWMode;
					break;
			}
			
			prcMode = modeValue;
			main.settings.setKey(Settings.prcMode, prcMode);
			
			if (currentPrcModeButton != null)
				currentPrcModeButton.active = false;
			currentPrcModeButton = modeBtn;
			currentPrcModeButton.active = true;
		}
		
		private function start():void
		{
			if (xmlFilePathError)
			{
				outputLogLine("Ошибка в пути к файлу с данными TrackChecker. Запуск невозможен.", COLOR_BAD);
				return;
			}
			
			if (!xmlFile.exists) 
			{
				outputLogLine("Файл с данными TrackChecker не найден", COLOR_BAD);
				return
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
			outputLogLine("Загрузка файла c данными TrackChecker");
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
				outputLogLine("Неверный формат файла с данными. Запуск отменён", COLOR_BAD);
				return;
			}
			
			maxID = maxIDDefault = uint(dataXml.@maxid);
			trace("MAX ID:", dataXml.@maxid);
			
			// Capture current date to use it in XML output
			currentDate = printDate();
			
			// Calling the right function based on current state of mode
			switch (prcMode) 
			{
				// Frontwinner provider
				case PRCMODE_FRONTWINNER:
					processFrontwinnerMode();
					break;
					
				// "Excel modes"
				// Canton 1 / 2, Shenzhen (SEO & CFF)
				//case PRCMODE_CANTON_WH_1:
				case PRCMODE_CANTON_WH_2:
				case PRCMODE_SHENZHEN:
					startLoadingExcelFile();
					break;
					
				default:
					throw new Error("Out of possible modes");
					break;
			}
		}
		
		private function startLoadingExcelFile():void
		{
			// Shenzhen mode only [!]
			if (prcMode == PRCMODE_SHENZHEN) 
			{
				if (xlDirPathError || !xlDir.exists || !xlDir.isDirectory)
				{
					outputLogLine("Ошибка в пути к папке с файлами Excel. Запуск невозможен", COLOR_BAD);
					return;
				}
				
				var dirContents:Array = xlDir.getDirectoryListing();
				
				if (dirContents.length == 0) 
				{
					outputLogLine("В указанной папке нет никаких файлов", COLOR_BAD);
					return;
				}
				
				var f:File;
				var xlFilesList:Vector.<File> = new Vector.<File>();
				for each (f in dirContents) 
				{
					if (f.nativePath.search(/.xlsx$/) != -1)
						xlFilesList.push(f);
				}
				
				if (xlFilesList.length == 0) 
				{
					outputLogLine("В указанной папке нет XLSX-файлов", COLOR_BAD);
					return;
				}
				
				var chosenXlFile:File;
				if (xlFilesList.length == 1) 
				{
					chosenXlFile = xlFilesList[0];
				}
				else 
				{
					var mostRecentFile:File;
					for each (f in xlFilesList) 
					{
						if (mostRecentFile == null)
							mostRecentFile = f;
						
						if (f.modificationDate.getTime() > mostRecentFile.modificationDate.getTime())
							mostRecentFile = f;
					}
					
					chosenXlFile = mostRecentFile;
				}
				
				outputLogLine("Файл <b>" + chosenXlFile.name + "</b> выбран из указанной папки", COLOR_SPECIAL);
			}
			
			// For rest Excel modes
			else 
			{
				// Excel file check
				if (xlFilePathError)
				{
					outputLogLine("Ошибка в пути к файлу Excel. Запуск невозможен.", COLOR_BAD);
					return;
				}
				
				chosenXlFile = xlFile;
			}
			
			// Load excel file
			xlLoader = new XLSXLoader();
			xlLoader.addEventListener(Event.COMPLETE, excelFileLoadingDone);
			xlLoader.load(chosenXlFile.nativePath);
			outputLogLine("Загрузка файла Excel");
		}
		
		private function excelFileLoadingDone(e:Event):void 
		{
			xlLoader.removeEventListener(Event.COMPLETE, excelFileLoadingDone);
			xlSheet = xlLoader.worksheet("Sheet1");
			
			// Mode fork (for excel modes)
			switch (prcMode) 
			{
				//case PRCMODE_CANTON_WH_1:
				case PRCMODE_CANTON_WH_2:
				{
					processCantonExcelModes1and2();
					break;
				}
				
				case PRCMODE_SHENZHEN:
				{
					processShenzhenMode();
					break;
				}
				
				default:
					throw new Error("Out of possible Excel modes");
					break;
			}
		}
		
		private function processCantonExcelModes1and2():void
		{
			if (xlSheet.getCellValue("A1").search(/Parcel List/i) == -1)
			{
				outputLogLine("Неверный формат таблицы для режима «Кантон»", COLOR_BAD);
				return;
			}
			
			// Local shortcuts for Canton modes
			var md1:Boolean = (prcMode == PRCMODE_CANTON_WH_1);
			var md2:Boolean = (prcMode == PRCMODE_CANTON_WH_2);
			
			outputLogLine("Чтение таблицы в формате Кантона", COLOR_SPECIAL);
			
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
			// Shortcut to source XML's root group
			var rootGroup:XML = dataXml.groups.(@id == 0)[0];
			
			// Iterate through our groups
			for each (var tscriptGrp:XML in groups.groups)
			{
				trace("TScript Group:", tscriptGrp.@desc);
				
				// Check if this group already exists
				var tcheckerExistingGroups:XMLList = rootGroup..groups.(trimSpaces(@desc) == tscriptGrp.@desc);
				
				// If exists
				if (tcheckerExistingGroups.length() > 0)
				{
					var tcheckerExistingGroup:XML;
					
					// Compare tracks in all existing groups to our current group
					tracksLoop:
					for each (var tscriptGrpTrack:XML in tscriptGrp.track)
					{
						trace("TScript Track:", tscriptGrpTrack.@track);
						
						// One operation on possible multiple dupes of our group
						for each (tcheckerExistingGroup in tcheckerExistingGroups)
						{
							var tcheckerExistingGroupTrackDups:XMLList = 
								tcheckerExistingGroup.track.(trimSpaces(@track) == tscriptGrpTrack.@track);
								
							if (tcheckerExistingGroupTrackDups.length() > 0)
							{
								main.logRed("Track Duplicate Found: " + tscriptGrpTrack.@desc + " " + tscriptGrpTrack.@track);
								existingTracksCount++;
								continue tracksLoop;
							}
							
							else
							{
								tcheckerExistingGroup.appendChild(tscriptGrpTrack);
							}
						}
					}
					
					// Stats
					existingGroupsCount++; // Our groups (there can be even more dupes of our group)
					
					// Skip to our next group
					continue;
				}
				
				// Add new groups to root group in Source XML
				rootGroup.appendChild(tscriptGrp);
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
				outputLogLine("Одинаковый прогон", COLOR_BAD);
				return;
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
			var headerDateAndOrderNumTemplate:RegExp = /^([\d-]+) ?(#\d+)(?= ?\$)/;
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
				
				// Retrieve [date] and [order number]
				reAr = l.match(headerDateAndOrderNumTemplate); // From Header line (it's currently being processed)
				currentTrackRecord.date = reAr[1]; // Date
				currentTrackRecord.orderNum = reAr[2]; // Order number with '#'
				
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
				
				// #SPECIAL: Netherlands only
				if (currentTrackRecord.country == "Netherlands") 
				{
					// Check the line before last on postal code
					var rea:Array = 
						trimSpaces(tmpRecordSourceLines[tmpRecordSourceLines.length-2])
							.match(postCodeTemplateForNetherlands);
					
					currentTrackRecord.postCode = rea[0];
				}
				
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
				newFrontwinnerGrp.appendChild(
					createXmlTrack(rec.name, rec.track, rec.country, rec.date, null, rec.orderNum, null, null, 
						rec.postCode));
			}
			
			var rootGroup:XML = dataXml.groups.(@id == 0)[0];
			
			var xmlQuery:XMLList;
			var frontwinnerGrp:XML;
			xmlQuery = rootGroup..groups.(trimSpaces(@desc).toLowerCase() == "Frontwinner".toLowerCase());
			if (xmlQuery.length() > 0)
			{
				frontwinnerGrp = xmlQuery[0];
				var newTracks:XMLList = newFrontwinnerGrp.track;
				for each (var t:XML in newTracks)
				{
					// Track Existence Check
					// If found duplicate > skip this track
					if (checkXmlTrackDuplicate(rootGroup, t.@track, t.@desc)) 
					{
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
			
			outputLogLine(
				"Обработано треков: " + tracksCount + 
				colorText(COLOR_TXLB_LIGHT_GREY, " · ") + 
				"Всего блоков: " + totalRecordsCount +
				(notTrackRecordsCount > 0 ? 
					colorText(COLOR_TXLB_LIGHT_GREY, " · ") + 
					colorText(COLOR_WARN, "Пропущенные блоки: " + notTrackRecordsCount) : "") +
				(invalidTrackRecordsCount > 0 ? 
					colorText(COLOR_TXLB_LIGHT_GREY, " · ") + 
					colorText(COLOR_BAD, "Ошибка в блоках: " + invalidTrackRecordsCount) : ""),
				
				null,
				true
			);
			
			outputLogLine(
				colorText(COLOR_SPECIAL, "Новые треки: " + (tracksCount - existingTracksCount).toString()) +
					
				(existingTracksCount > 0 ? 
					colorText(COLOR_TXLB_LIGHT_GREY, " · ") + colorText(COLOR_WARN, "Дубли: " + existingTracksCount) : 
						colorText(COLOR_TXLB_LIGHT_GREY, " · ") + colorText(COLOR_TXLB_LIGHT_GREY, "Дублей не найдено")),
				
					null,
					true
			); 
			
			if (tracksCount == existingTracksCount && tracksCount > 0)
			{
				outputLogLine("Одинаковый прогон", COLOR_BAD);
				return;
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
		}
		
		private function processShenzhenMode():void 
		{
			/*
			Алгоритм
			> Validation checks
			> Determine columns
			> Parse excel
			> Fill XML
			> Show stats
			> Write to output file
			> Write XML back to file
			*/
			
			/**
			 * General validation
			 * ================================================================================
			 */
			
			if (xlSheet.cols <= 1) 
			{
				outputLogLine("Неверный формат таблицы для режима «Шэньчжэнь»", COLOR_BAD);
				return;
			}
			
			// User defined date
			var userDefinedDateStr:String = 
				trimSpaces(ui.nsDateDay.textField.text) + "." + 
				trimSpaces(ui.nsDateMonth.textField.text) + "." +
				trimSpaces(ui.nsDateYear.textField.text);
				
			const userDefinedDatePattern:RegExp = /^(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|1[012])\.(\d{4})$/;
			
			if (userDefinedDateStr == "") 
			{
				userDefinedDate = null;
				outputLogLine("Дата для события «Added» не указана", COLOR_BAD);
				return;
			}
			
			else if (userDefinedDateStr.search(userDefinedDatePattern) == -1) 
			{
				userDefinedDate = null;
				outputLogLine("Указана неправильная дата", COLOR_BAD);
				return;
			}
			
			else 
			{
				var reAr:Array = userDefinedDateStr.match(userDefinedDatePattern);
				userDefinedDate = new Date(int(reAr[3]), int(reAr[2]) - 1, int(reAr[1]));
				outputLogLine("Дата для события «Added»: " + userDefinedDateStr, COLOR_SPECIAL);
			}
			
			// ================================================================================
			
			var i:int;
			
			outputLogLine("Чтение таблицы в формате Шэньчжэня", COLOR_SPECIAL);
			
			/**
			 * Determine columns
			 * ================================================================================
			 */
			
			const tableHeaderRow:int = 1;
			const trackColHeaderPattern:RegExp = /^Tracking no/i;
			const nameColHeaderPattern:RegExp = /^Buyer Fullname/i;
			const countryColHeaderPattern:RegExp = /^Buyer Country/i;
			const postCodeColHeaderPattern:RegExp = /^Buyer Zip/i;
			const skuColHeaderPattern:RegExp = /^(SKU|Custom Label)/i;
			const weightColHeaderPattern:RegExp = /^Chargeable weight/i;
			const packageOrderNumColHeaderPattern:RegExp = /^Order num/i;
			const totalCostColHeaderPattern:RegExp = /^Total amount/i;
			const quantityColHeaderPattern:RegExp = /^Quantity/i;
			
			var trackCol:String;
			var nameCol:String;
			var countryCol:String;
			var postCodeCol:String;
			var skuCol:String;
			var weightCol:String;
			var packageOrderNumCol:String;
			var totalCostCol:String;
			var quantityCol:String;
			
			if (xlSheet.cols > xlColLetters.length) 
			{
				outputLogLine("Количество колонок в таблице выходит за пределы разумного", COLOR_BAD);
				return;
			}
			
			var headerCellVal:String;
			var len:int = xlSheet.cols;
			for (i = 0; i < len; i++) 
			{	
				headerCellVal = trimSpaces(xlSheet.getCellValue(xlColLetters[i] + tableHeaderRow));
				
				// Track col header
				if (trackCol == null && headerCellVal.search(trackColHeaderPattern) != -1)
				{
					trackCol = xlColLetters[i];
				}
				
				// Name col header
				if (nameCol == null && headerCellVal.search(nameColHeaderPattern) != -1) 
				{
					nameCol = xlColLetters[i];
				}
				
				// Country col header
				if (countryCol == null && headerCellVal.search(countryColHeaderPattern) != -1) 
				{
					countryCol = xlColLetters[i];
				}
				
				// Postcode col header
				if (postCodeCol == null && headerCellVal.search(postCodeColHeaderPattern) != -1) 
				{
					postCodeCol = xlColLetters[i];
				}
				
				// SKU col header
				if (skuCol == null && headerCellVal.search(skuColHeaderPattern) != -1) 
				{
					skuCol = xlColLetters[i];
				}
				
				// Weight col header
				if (weightCol == null && headerCellVal.search(weightColHeaderPattern) != -1) 
				{
					weightCol = xlColLetters[i];
				}
				
				// Order num col header
				if (packageOrderNumCol == null && headerCellVal.search(packageOrderNumColHeaderPattern) != -1) 
				{
					packageOrderNumCol = xlColLetters[i];
				}
				
				// Total cost col header
				if (totalCostCol == null && headerCellVal.search(totalCostColHeaderPattern) != -1) 
				{
					totalCostCol = xlColLetters[i];
				}
				
				// Quantity col header
				if (quantityCol == null && headerCellVal.search(quantityColHeaderPattern) != -1) 
				{
					quantityCol = xlColLetters[i];
				}
				
				// Final check
				if (trackCol != null && 
					nameCol != null &&
					countryCol != null &&
					postCodeCol != null &&
					skuCol != null &&
					weightCol != null &&
					packageOrderNumCol != null &&
					totalCostCol != null &&
					quantityCol != null) 
				{
					// All columns determined
					break;
				}
				else if (i == len-1)
				{
					outputLogLine("Не все требуемые колонки определены", COLOR_BAD);
					return;	
				}
			}
			
			/*
			outputLogLine(
				"Колонка с треками: " + trackCol +
				", Имя покупателя: " + nameCol +
				", Страна: " + countryCol +
				", Индекс: " + postCodeCol +
				", SKU: " + skuCol +
				", Вес посылки: " + weightCol +
				", Номер заказа: " + packageOrderNumCol +
				", Стоимость: " + totalCostCol +
				", Количество товара: " + quantityCol
			);
			*/
			
			/**
			 * Parse table
			 * ================================================================================
			 */
						
			var row:int;
			var packages:Vector.<ShenzhenPackage> = new Vector.<ShenzhenPackage>();
			var currentPackage:ShenzhenPackage;
			 
			var trackColVal:String;
			var nameColVal:String;
			var countryColVal:String;
			var postCodeColVal:String;
			var skuColVal:String;
			var weightColVal:String;
			var packageOrderNumColVal:String;
			var totalCostColVal:String;
			var quantityColVal:String;
			
			function initPackage():void
			{
				currentPackage = new ShenzhenPackage();
				
				currentPackage.track = trackColVal;
				currentPackage.buyerName = nameColVal;
				currentPackage.buyerCountry = countryColVal;
				currentPackage.buyerPostCode = postCodeColVal;
				currentPackage.itemsList = new Vector.<String>();
				currentPackage.itemsQuantityList = new Dictionary();
				
				// SKU
				if (skuColVal != null)
					currentPackage.itemsList.push(skuColVal);
				else
					outputLogLine("Пустой SKU на строке " + row, COLOR_WARN);
				
				// Weight
				if (weightColVal != null)
					currentPackage.weight = weightColVal;
				else
					outputLogLine("Пустой вес на строке " + row, COLOR_WARN);
				
				// Order num
				if (packageOrderNumColVal != null)
					currentPackage.packageOrderNum = packageOrderNumColVal;
				else
					outputLogLine("Пустой номер заказа на строке " + row, COLOR_WARN);
					
				// Total cost
				if (totalCostColVal != null)
					currentPackage.totalCost = totalCostColVal;
				else
					outputLogLine("Пустая стоимость на строке " + row, COLOR_WARN);
					
				// Quantity
				if (quantityColVal != null && skuColVal != null)
				{
					currentPackage.singleItemQuantity = uint(quantityColVal);
					
					if (skuColVal != null)
						currentPackage.itemsQuantityList[skuColVal] = uint(quantityColVal);
				}
				else
				{
					outputLogLine("Проблема с количеством товара на строке " + row, COLOR_WARN);
				}
			}
			
			function finishPackage():void 
			{
				packages.push(currentPackage);
				currentPackage = null;
				tracksCount++;
			}
			
			// Reset stats
			tracksCount = existingTracksCount = 0;
			
			row = 1;
			
			while (true) 
			{
				row++;
				trackColVal = trimSpaces(xlSheet.getCellValue(trackCol + row));
				nameColVal = trimSpaces(xlSheet.getCellValue(nameCol + row));
				countryColVal = trimSpaces(xlSheet.getCellValue(countryCol + row));
				postCodeColVal = trimSpaces(xlSheet.getCellValue(postCodeCol + row));
				skuColVal = trimSpaces(xlSheet.getCellValue(skuCol + row));
				weightColVal = trimSpaces(xlSheet.getCellValue(weightCol + row));
				packageOrderNumColVal = trimSpaces(xlSheet.getCellValue(packageOrderNumCol + row));
				totalCostColVal = trimSpaces(xlSheet.getCellValue(totalCostCol + row));
				quantityColVal = trimSpaces(xlSheet.getCellValue(quantityCol + row));
				
				// Determine line type
				// · Header line
				if (trackColVal != "" && nameColVal != "" && countryColVal != "") 
				{
					if (currentPackage != null)
						finishPackage();
						
					initPackage();
					continue;
				}
				
				// · Secondary valid line
				else if (skuColVal != "" && trackColVal == "" && nameColVal == "" && countryColVal == "") 
				{
					if (currentPackage != null)
					{
						currentPackage.itemsList.push(skuColVal);
						currentPackage.itemsQuantityList[skuColVal] = uint(quantityColVal);
						continue;
					}
					else 
					{
						outputLogLine("Осиротевшая строка " + row + " проигнорирована", COLOR_WARN);
						continue;
					}
				}
				
				// · Invalid line
				else 
				{
					// Empty line
					if (trackColVal == "" && nameColVal == "" && countryColVal == "" && skuColVal == "" && weightColVal == "") 
					{
						if (currentPackage != null)	
							finishPackage();
						
						if (row >= xlSheet.rows)
							// Finish parsing
							break;
					}
					
					outputLogLine("Проблема со строкой " + row + ". Она проигнорирована", COLOR_WARN);
					continue;
				}
			}
			
			trace("Parsing done");
			
			/**
			 * Fill XML
			 * ================================================================================
			 */
			
			var rootGroup:XML = dataXml.groups.(@id == 0)[0];
			var xmlQuery:XMLList;
			var seoFolder:XML;
			var pkg:ShenzhenPackage;
			var commentWithSkuList:String = null;
			
			userDefinedDateStr = userDefinedDate != null ? printDate(true, userDefinedDate) : currentDate;
			
			xmlQuery = rootGroup.groups.(@desc == "SEO");
			if (xmlQuery.length() > 0) 
			{
				seoFolder = xmlQuery[0];
			}
			else 
			{
				seoFolder = createXmlGroup("SEO");
				rootGroup.prependChild(seoFolder);
			}
			
			for each (pkg in packages) 
			{
				if (checkXmlTrackDuplicate(dataXml, pkg.track, pkg.buyerName)) 
				{
					existingTracksCount++;
					continue;
				}
				
				// SKU list in tracks' comment
				if (pkg.itemsList.length > 0) 
				{
					commentWithSkuList = "";
					
					for each (var skuEntry:String in pkg.itemsList) 
					{
						if (uint(pkg.itemsQuantityList[skuEntry]) > 1) 
							commentWithSkuList += skuEntry + " (x" + String(pkg.itemsQuantityList[skuEntry]) + ")";
						else 
							commentWithSkuList += skuEntry;
						
						// New line symbol
						commentWithSkuList += skuEntry != pkg.itemsList[pkg.itemsList.length-1] ? "\n" : "";
					}
				}
				
				seoFolder.appendChild(
					createXmlTrack(pkg.buyerName, pkg.track, pkg.buyerCountry, 
						userDefinedDateStr, commentWithSkuList, pkg.packageOrderNum, 
						pkg.weight, pkg.totalCost, pkg.buyerPostCode)
				);
			}
			
			trace("Filling done");
			
			// Write result of script work to output.xml
			var output:XML = <track-script-output/>;
			output.appendChild(seoFolder);
			output.@date = getFormattedDate("dd.MM.yyyy HH:mm:ss");
			output.@version = main.version;
			output.@format = "Shenzhen";
			output.@excel = ui.tfXlFile.text;
			writeToOutputFile(output);
			
			if (tracksCount == existingTracksCount && tracksCount > 0)
			{
				outputLogLine("Дубли треков: " + existingTracksCount, COLOR_WARN);
				outputLogLine("Одинаковый прогон — 100% дублей треков", COLOR_BAD);
				return;
			}
			
			/**
			 * Weight stat
			 * ================================================================================
			 */
			
			// Check file presence
			checkWeightStatFile(); 
			 
			// Local stats
			var weightStatProducts:uint = 0;
			var skippedWeightStatProducts:uint = 0;
			
			fst = new FileStream();
			fst.open(weightStatFile, FileMode.READ);
			var weightStatXmlStr:String = fst.readUTFBytes(fst.bytesAvailable);
			fst.close();
			
			if (weightStatXmlStr.length > 0) 
			{
				weightStatXml = new XML(weightStatXmlStr);
				
				var p:XML; // Shortcut for product record (entry) in weight stat XML
				var w:XML; // Shortcut for weight record (entry)
				
				function processWeightEntry(appendOverflow:Boolean = false):void 
				{
					// Create new entry
					w = <w/>;
					w.@v = pkg.weight;
					w.@pon = pkg.packageOrderNum;
					p.appendChild(w);
					
					// If maximum entries in the list > also delete the first item in the list
					if (appendOverflow)
						delete p.children()[0];
				}
				
				for each (pkg in packages) 
				{
					if (pkg.itemsList.length == 1 && pkg.weight != null && pkg.weight != "") 
					{
						if (pkg.singleItemQuantity > 1)
							pkg.weight = (Number(pkg.weight) / pkg.singleItemQuantity).toString();
						
						xmlQuery = weightStatXml.p.(@sku == pkg.itemsList[0]);
						if (xmlQuery.length() > 0) 
						{
							p = xmlQuery[0];
							
							// Check weight entry existence in stat by package order num
							xmlQuery = p.w.(attribute("pon") == pkg.packageOrderNum);
							if (xmlQuery.length() > 0) 
							{
								// If entry with such order num already exists > skip this package for adding to stat
								skippedWeightStatProducts++;
								continue;
							}
							
							// If no entry presence by order num (above) > Add weight entry to product node in stat
							if (p.children().length() >= WEIGHT_STAT_MAX_RECORD_COUNT) 
							{
								processWeightEntry(true);
							}
							else 
							{
								processWeightEntry();
							}
						}
						else 
						{
							p = <p/>;
							p.@sku = pkg.itemsList[0];
							processWeightEntry();
							weightStatXml.appendChild(p);
						}
						
						weightStatProducts++;
					}
				}
				
				XML.prettyPrinting = true;
				XML.prettyIndent = 4;
				
				fst = new FileStream();
				fst.open(weightStatFile, FileMode.WRITE);
				fst.writeUTFBytes(weightStatXml.toXMLString());
				fst.close();
			}
			
			/**
			 * Show Stats
			 * ================================================================================
			 */
			outputLogLine("Всего обработано посылок из экселя: " + tracksCount);
			
			outputLogLine(
				colorText(COLOR_SPECIAL, "Новые треки: " + (tracksCount - existingTracksCount).toString() /*+
					" из " + tracksCount*/) +
					/*colorText(COLOR_TXLB_LIGHT_GREY, 
						" (" + Math.round(
								Number(((tracksCount - existingTracksCount) / tracksCount) * 100)
							).toString() + "%)") +*/
					
				(existingTracksCount > 0 ? 
					colorText(COLOR_TXLB_LIGHT_GREY, " · ") + colorText(COLOR_WARN, "Дубли: " + existingTracksCount) : 
						colorText(COLOR_TXLB_LIGHT_GREY, " · ") + colorText(COLOR_TXLB_LIGHT_GREY, "Дублей не найдено")),
				
					null,
					true
			);
				
			if (weightStatProducts > 0 || skippedWeightStatProducts > 0)
			{
				if (weightStatProducts == 0 && skippedWeightStatProducts > 0)
				
					outputLogLine(
						colorText(COLOR_WARN, "Вес не добавлен в статистику: " + skippedWeightStatProducts) /*+
							colorText(COLOR_TXLB_LIGHT_GREY, " (100%)")*/, 
							
						null,
						true
					);
					
				else if (weightStatProducts > 0)
				
					outputLogLine(
						colorText(COLOR_SPECIAL, "Вес добавлен в статистику: " + weightStatProducts /*+ 
							" из " + (weightStatProducts + skippedWeightStatProducts).toString()*/) +
							/*colorText(COLOR_TXLB_LIGHT_GREY, 
								" (" + Math.round(
										Number((weightStatProducts / (weightStatProducts + skippedWeightStatProducts)) * 100)
									).toString() + "%)") +*/
									
						(skippedWeightStatProducts > 0 ? colorText(COLOR_TXLB_LIGHT_GREY, " · ") + 
							colorText(COLOR_WARN, "Не добавлен: " + skippedWeightStatProducts) : ""),
							
						null,
						true
					);
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
			
			outputLogLine("Готово. Добавлено в TrackChecker", COLOR_SUCCESS);
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
		
		private function createXmlTrack(
			name:String, track:String, country:String,
			eventSpecialDate:String = null, comment:String = null, orderNum:String = null, 
			weight:String = null, totalCost:String = null, postCode:String = null):XML
		{
			var trackOutput:String; // For track value modifications
			
			trackOutput = track; // Standart track value output
			
			// ====================
			// ~ TRACK VALUE MODS ~
			// ====================
			
			// #SPECIAL: Netherlands
			if (country.search(/Netherlands/i) != -1 &&
				track.search(/CN$/i) != -1)
			{
				if (postCode == null)
					outputLogLine("Проблема с индексом в треке для NL " + track + " (" + name + ")");
				else
					trackOutput = track + "/NL/" + clearSpaces(postCode); // Mod
			}
			
			// =========================
			// ~ TRACK ENTRY FORMATION ~
			// =========================
			
			// New track
			var xmlTrack:XML = <track/>;
			
			xmlTrack.@id = ++maxID; 
			xmlTrack.@desc = name;
			xmlTrack.@crdt = currentDate;
			xmlTrack.@track = trackOutput;
			
			// Comment
			if (comment != null)
				xmlTrack.@comm = comment;
			
			// ============
			// ~ SERVICES ~
			// ============
			
			var xmlTrackServs:XML = <servs/>;
			var xmlServ:XML;
			var servAlias:String;
			
			xmlTrackServs.@id = ++maxID;
			xmlTrackServs.@crdt = currentDate;
			
			// Regular standart services
			var standartServAliases:Array = ["gtt"]; // Prev: 'china_ems'
			
			// #SPECIAL: Special standart services
			if (track.search(/HK$/i) != -1) 
			{
				standartServAliases = ["hkpost"];
			}
			else if (track.search(/EE$/i) != -1)
			{
				standartServAliases = ["ee_post11"];
			}
			else if (track.search(/CH$/i) != -1)
			{
				standartServAliases = ["swi"];
			}
			else if (track.search(/NL$/i) != -1)
			{
				standartServAliases = ["nl_post2"];
			}
			else if (track.search(/DE$/i) != -1)
			{
				standartServAliases = ["dhl_ec_asia"];
			}
			
			for each (servAlias in standartServAliases)
			{
				xmlServ = <serv/>;
				xmlServ.@id = ++maxID;
				xmlServ.@crdt = currentDate;
				xmlServ.@serv = servAlias;
				xmlServ.@selected = 1;
				xmlTrackServs.appendChild(xmlServ);
			}
			
			// Special services based on country
			var specialServs:Array;
			
			// #SPECIAL-CASE: UK service in Shenzhen mode
			if (prcMode == PRCMODE_SHENZHEN && country.search(/United Kingdom/i) != -1) 
			{
				specialServs = ["uk_yodel"];
			}
			
			// Regular behavior: take special services based on country from user file
			if (specialServs == null)
				specialServs = getCntServices(country);
			
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
			
			xmlTrack.appendChild(xmlTrackServs);
			
			// ==========
			// ~ EVENTS ~
			// ==========
			
			var eventDateCRDT:String = eventSpecialDate == null ? currentDate : eventSpecialDate + "T00:00:00";
			var eventDateUDT:String = eventSpecialDate == null ? printDate(true) : eventSpecialDate;
			
			// Description events
			if (prcMode == PRCMODE_FRONTWINNER || prcMode == PRCMODE_SHENZHEN) 
			{
				var xmlEventDesc:XML;
				var i:int; // For loops
				
				// 4 iterations = 2 events (Frontwinner) and 4 events (Shenzhen)
				for (i = 1; i < 4+1; i++) 
				{	
					if (i > 2 && prcMode == PRCMODE_FRONTWINNER)
						continue;
					
					xmlEventDesc = <event/>;
					xmlEventDesc.@id = ++maxID;
					xmlEventDesc.@crdt = eventDateCRDT
					
					if (prcMode == PRCMODE_FRONTWINNER) 
					{
						xmlEventDesc.@desc = i == 1 ? "Страна: " + country : (i == 2 ? "Заказ: " + orderNum : "");
					}
					else if (prcMode == PRCMODE_SHENZHEN)
					{
						if (i == 1) xmlEventDesc.@desc = "Страна: " + country;
						if (i == 2) xmlEventDesc.@desc = "Заказ: " + orderNum;
						if (i == 3)
							xmlEventDesc.@desc = "Вес: " + ((Number(weight) < 1) ? (Number(weight) * 1000).toString() + " г" : weight + " кг");
						if (i == 4) xmlEventDesc.@desc = "Стоимость: " + Number(totalCost).toString();
					}
					
					xmlEventDesc.@udt = eventDateUDT
					xmlEventDesc.@auto = 1;
					xmlEventDesc.@info = 1;
					
					xmlTrack.appendChild(xmlEventDesc);
				}
			}
			
			// Event "Added". Appended to every track
			var xmlEventAdded:XML = <event/>;
			xmlEventAdded.@id = ++maxID;
			xmlEventAdded.@crdt = eventDateCRDT
			xmlEventAdded.@desc = "Added";
			xmlEventAdded.@udt = eventDateUDT
			
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
		
		private function checkXmlTrackDuplicate(xmlGroupToExplore:XML, trackVal:String, trackDesc:String):Boolean 
		{
			var xmlQuery:XMLList = 
				xmlGroupToExplore..track.(trimSpaces(@desc).indexOf(trackDesc) != -1 && trimSpaces(@track).indexOf(trackVal) != -1);
				
			if (xmlQuery.length() > 0) 
				return true;
			
			return false;
		}
		
		private function searchExcelColumnByHeader(xlSheet:Worksheet, headerPattern:RegExp, headerRow:int, maxColumn:String):String 
		{
			return "";
		}
		
		private function searchExcelColumnByValue(xlSheet:Worksheet, valuePattern:RegExp, maxColumn:String):String 
		{
			return "";
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
		
		private function checkWeightStatFile():void 
		{
			weightStatFile = File.applicationStorageDirectory.resolvePath("weight-stat.xml");
			if (!weightStatFile.exists || weightStatFile.size == 0) 
			{
				XML.prettyPrinting = true;
				XML.prettyIndent = 4;
				var defWeightStatXml:XML = <weight-stat/>;
				
				fst = new FileStream();
				fst.open(weightStatFile, FileMode.WRITE);
				fst.writeUTFBytes(defWeightStatXml.toXMLString());
				fst.close();
			}
		}
		
		private function printDate(noTime:Boolean = false, customDate:Date = null):String
		{
			var d:Date = customDate != null ? customDate : new Date();
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
		
		private function clearSpaces(str:String):String
		{
			return str.replace(/\s+/, "");
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