package {

	import flash.events.Event;

	/**
	 * ...
	 * @author Tim Yusupov
	 */
	public class SettingEvent extends Event {

		public static const VALUE_CHANGED:String = "valueChanged";

		private var $settingName:String;
		private var $newValue:*;

		public function SettingEvent(type:String, settingName:String, newValue:*, bubbles:Boolean = false, cancelable:Boolean = false):void {
			super(type, bubbles, cancelable);
			$settingName = settingName;
			$newValue = newValue;
		}

		public override function clone():Event {
			return new SettingEvent(type, $settingName, $newValue, bubbles, cancelable);
		}

		public override function toString():String {
			return formatToString("SettingEvent", "type", "bubbles", "cancelable", "eventPhase");
		}

		public function get settingName():String {
			return $settingName;
		}

		public function get newValue():* {
			return $newValue;
		}

	}

}