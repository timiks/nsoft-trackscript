package entities 
{
		
	/**
	 * ...
	 * @author Tim Yusupov
	 */
	public class ShenzhenPackage 
	{
		private var $track:String;
		private var $buyerName:String;
		private var $buyerCountry:String;
		private var $weight:String;
		private var $itemsList:Vector.<String>;
		
		public function ShenzhenPackage():void 
		{
			
		}
		
		public function get track():String 
		{
			return $track;
		}
		
		public function set track(value:String):void 
		{
			$track = value;
		}
		
		public function get buyerName():String 
		{
			return $buyerName;
		}
		
		public function set buyerName(value:String):void 
		{
			$buyerName = value;
		}
		
		public function get buyerCountry():String 
		{
			return $buyerCountry;
		}
		
		public function set buyerCountry(value:String):void 
		{
			$buyerCountry = value;
		}
		
		public function get weight():String 
		{
			return $weight;
		}
		
		public function set weight(value:String):void 
		{
			$weight = value;
		}
		
		public function get itemsList():Vector.<String> 
		{
			return $itemsList;
		}
		
		public function set itemsList(value:Vector.<String>):void 
		{
			$itemsList = value;
		}
	}
}