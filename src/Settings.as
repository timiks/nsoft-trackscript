package {

	import flash.events.Event;
	import flash.events.EventDispatcher;
	import flash.events.TimerEvent;
	import flash.filesystem.File;
	import flash.filesystem.FileMode;
	import flash.filesystem.FileStream;
	import flash.net.FileReference;
	import flash.utils.ByteArray;
	import flash.utils.Timer;

	/**
	 * ...
	 * @author Tim Yusupov
	 */
	public class Settings {

		public function Settings():void {}

		public static const sourceExcelFile:String = "sourceExcelFile";
		public static const trackCheckerDataFile:String = "trackCheckerDataFile";
		public static const winPos:String = "winPos";

		private var main:Main;
		private var allSets:Object;
		private var sets:Object;
		private var settingsFile:File;
		private var fstream:FileStream;
		private var tmrSaveDelay:Timer;

		private var $eventDsp:EventDispatcher;

		private var loaded:Boolean = false;

		public function load():void {

			main = Main.ins;
			$eventDsp = new EventDispatcher();

			allSets = {}; 							// Name / Data type / Default value / Version introduced in
			allSets[sourceExcelFile] 				= new Setting(sourceExcelFile, String, "", "1.0");
			allSets[trackCheckerDataFile] 			= new Setting(trackCheckerDataFile, String, "", "1.0");
			allSets[winPos]							= new Setting(winPos, String, "374:163", "1.0");

			sets = {};
			settingsFile = File.applicationStorageDirectory.resolvePath("settings.json");
			fstream = new FileStream();

			if (!settingsFile.exists) {

				trace("Settings File not found");
				initDefSettingsAndCreateFile();
				loaded = true;

			} else {

				trace("Settings File found: " + settingsFile.nativePath);

				fstream.open(settingsFile, FileMode.READ);
				var jsonString:String = fstream.readUTFBytes(fstream.bytesAvailable);
				var retObj:Object = parseJSON(jsonString);
				fstream.close();

				if (retObj == null) { // error
					// Back up corrupted file
					var date:Date = new Date();
					var dateTime:String = date.getDay() + "." + date.getMonth() + "." +
					date.getFullYear() + "-" + date.getHours() + "." + date.getMinutes() + "." + date.getSeconds();
					var badFileBackUp:FileReference = File.applicationStorageDirectory.resolvePath("settings-"+dateTime+".bak");
					settingsFile.copyTo(badFileBackUp, true);
					settingsFile.deleteFile();

					// [To-Do Here ↓]: init default
					initDefSettingsAndCreateFile();

				} else {
					sets = retObj;
					retObj = null;
					validateSettings();
				}

				loaded = true;

			}

			/*
			fstream.addEventListener(IOErrorEvent.IO_ERROR, ioError);
			fstream.addEventListener(ProgressEvent.PROGRESS, streamProgress);
			fstream.addEventListener(OutputProgressEvent.OUTPUT_PROGRESS, outputProgress);
			*/

		}

		public function setKey(key:String, value:*):void {
			sets[key] = value;

			$eventDsp.dispatchEvent(new SettingEvent(SettingEvent.VALUE_CHANGED, key, value));

			if (tmrSaveDelay == null) {
				tmrSaveDelay = new Timer(3000, 1);
				tmrSaveDelay.addEventListener(TimerEvent.TIMER, saveOnTimer);
				tmrSaveDelay.start();
			}
		}

		public function getKey(key:String):* {
			return sets[key] != null ? sets[key] : null;
		}

		private function saveOnTimer(e:TimerEvent):void {
			saveFile();
			tmrSaveDelay.removeEventListener(TimerEvent.TIMER, saveOnTimer);
			tmrSaveDelay = null;
		}

		private function initDefSettingsAndCreateFile(createFile:Boolean = true):void {
			// Default Settings
			for each (var setting:Setting in allSets) {
				sets[setting.name] = setting.defValue;
			}

			if (!createFile) return;

			// Pack to JSON and create Settings File
			fstream.open(settingsFile, FileMode.WRITE);
			fstream.writeUTFBytes(packJSON(sets));
			fstream.close();
		}

		public function saveFile():void {
			fstream.open(settingsFile, FileMode.WRITE);
			fstream.writeUTFBytes(packJSON(sets));
			fstream.close();
			main.logRed("Settings File Saved");
		}

		private function packJSON(obj:Object):String {
			return JSON.stringify(obj, null, 4); // Pretty JSON
		}

		private function parseJSON(str:String):Object {
			var jsonObj:Object;

			try {

				jsonObj = JSON.parse(str);
				return jsonObj;

			} catch (err:Error) {

				main.logRed("JSON PARSE ERROR");
				return null;

			}

			return null;
		}

		private function validateSettings():void {
			var changed:Boolean = false;

			// Redundant settings and value type correctness checks
			for (var settingName:String in sets) {
				if (allSets[settingName] == null) {
					delete sets[settingName];
					changed = true;
					continue;
				}

				if (!(sets[settingName] is (allSets[settingName] as Setting).dataType)) {
					sets[settingName] = (allSets[settingName] as Setting).defValue;
					changed = true;
				}
			}

			// Absence of settings check
			for each (var setting:Setting in allSets) {
				if (sets[setting.name] == null) {
					sets[setting.name] = setting.defValue;
					changed = true;
				}
			}

			if (changed) saveFile();
		}

		// PROPERTY: eventDsp
		// ================================================================================

		public function get eventDsp():EventDispatcher {
			return $eventDsp;
		}

	}

}

class Setting {

	public var name:String;
	public var defValue:*;
	public var dataType:Class;
	public var versionIntroducedIn:String;

	public function Setting(name:String, dataType:Class, defValue:*, versionIntroducedIn:String = null):void {
		this.name = name;
		this.dataType = dataType;
		this.defValue = defValue;
		this.versionIntroducedIn = versionIntroducedIn;
	}

}