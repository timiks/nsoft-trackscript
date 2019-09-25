package
{
	import flash.desktop.NativeApplication;
	import flash.display.Sprite;
	import flash.events.Event;
	import flash.events.InvokeEvent;
	
	/**
	 * ...
	 * @author Tim Yusupov
	 */
	public class Main extends Sprite
	{
		private static var $ins:Main;
		
		// Version
		private const $version:int 				= 9;
		private const $versionService:int 		= 4;
		private const $betaVersion:Boolean 		= false;
		private const $betaVersionNumber:int 	= 0;
		private const bugs:Boolean 				= false;
		
		// Functional Members
		private var $settings:Settings;
		private var $mState:TrackScript;
		
		private var inited:Boolean;
		private var args:Array;
		
		public function Main():void
		{
			args = [];
			NativeApplication.nativeApplication.addEventListener(InvokeEvent.INVOKE, appInvoke);
		}
		
		private function appInvoke(e:InvokeEvent):void
		{
			args = e.arguments;
			/*
			if (inited)
			{
				
				if (args[0] == "")
				{
					
				}
				
				return;
			}
			*/
			stage ? init() : addEventListener(Event.ADDED_TO_STAGE, init);
		}
		
		private function init(e:Event = null):void
		{
			removeEventListener(Event.ADDED_TO_STAGE, init);
			
			$ins = this;
			
			/**
			 * Initialization
			 * ================================================================================
			 */
			// Settings
			$settings = new Settings();
			$settings.load();
			
			// Main State; Add it to DisplayList
			$mState = new TrackScript();
			addChild($mState);
			
			inited = true;
		}
		
		public function logRed(str:String):void
		{
			trace("3:" + str);
		}
		
		public function exitApp():void
		{
			trace("");
			trace("App is terminating");
			NativeApplication.nativeApplication.dispatchEvent(new Event(Event.EXITING));
			settings.saveFile();
			NativeApplication.nativeApplication.exit();
		}
		
		// STATIC PROPERTY: ins
		// ================================================================================
		
		public static function get ins():Main
		{
			if ($ins == null)
				throw new Error("Accessing Main while it isn't initialized");
			return $ins;
		}
		
		// PROPERTY: version
		// ================================================================================
		
		public function get version():String
		{
			var vr:String = String($version);
			if ($versionService > 0) vr += "." + String($versionService);
			if ($betaVersion)
			{
				vr += " β";
				if ($betaVersionNumber > 0) vr += String($betaVersionNumber);
			}
			return vr;
		}
		
		// PROPERTY: settings
		// ================================================================================
		
		public function get settings():Settings
		{
			return $settings;
		}
		
		// PROPERTY: windowUI
		// ================================================================================
		
		public function get mState():TrackScript
		{
			return $mState;
		}
	}
}