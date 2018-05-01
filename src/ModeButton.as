package 
{
	import flash.display.Graphics;
	import flash.display.MovieClip;
	import flash.display.Shape;
	import flash.events.Event;
	import flash.events.EventDispatcher;
	import flash.events.MouseEvent;
	import flash.text.TextField;
	import flash.text.TextFormat;
		
	/**
	 * ...
	 * @author Tim Yusupov
	 */
	public class ModeButton extends EventDispatcher
	{
		private var mc:MovieClip;
		private var tf:TextField;
		private var backRec:Shape;
		
		private var $active:Boolean;
		private var $linkedMode:int;
		
		public function ModeButton(btnMc:MovieClip, linkedMode:int, active:Boolean = false):void 
		{
			this.mc = btnMc;
			this.tf = mc.getChildByName("tf") as TextField;
			$linkedMode = linkedMode;
			
			backRec = new Shape();
			mc.addChildAt(backRec, 0);
			
			this.active = active;
			
			mc.buttonMode = true;
			mc.useHandCursor = false;
			tf.mouseEnabled = false;
			
			mc.addEventListener(MouseEvent.ROLL_OVER, onRollOver);
			mc.addEventListener(MouseEvent.ROLL_OUT, onRollOut);
			mc.addEventListener(MouseEvent.CLICK, onClick);
		}
		
		private function onClick(e:MouseEvent):void 
		{
			dispatchEvent(new Event("click")); // Custom click event
		}
		
		private function onRollOver(e:MouseEvent):void 
		{
			var txFrm:TextFormat = tf.getTextFormat();
			txFrm.color = 0xCC171C; // Red
			tf.setTextFormat(txFrm);
		}
		
		private function onRollOut(e:MouseEvent):void 
		{
			if (active) return;
			var txFrm:TextFormat = tf.getTextFormat();
			txFrm.color = 0;
			tf.setTextFormat(txFrm);
		}
		
		public function get active():Boolean 
		{
			return $active;
		}
		
		public function set active(value:Boolean):void 
		{
			$active = value;
			
			var rg:Graphics = backRec.graphics;
			var txFrm:TextFormat;
			
			if (value == true) 
			{
				rg.clear();
				rg.beginFill(0xFFFFFF, 0.5);
				rg.drawRect(0, 0, tf.width, tf.height);
				rg.endFill();
				rg.lineStyle(1, 0x0075BF);
				rg.drawRect(0, 0, tf.width, tf.height);
				
				txFrm = tf.getTextFormat();
				txFrm.color = 0x0075BF;
				tf.setTextFormat(txFrm);
				
				mc.mouseEnabled = false;
			}
			else
			{
				rg.clear();
				rg.beginFill(0xFFFFFF, 0);
				rg.drawRect(0, 0, tf.width, tf.height);
				rg.endFill();
				
				txFrm = tf.getTextFormat();
				txFrm.color = 0;
				tf.setTextFormat(txFrm);
				
				mc.mouseEnabled = true;
			}
		}
		
		public function get linkedMode():int 
		{
			return $linkedMode;
		}
	}
}