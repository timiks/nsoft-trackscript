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
		private var $buyerPostCode:String;
		private var $weight:String;
		private var $packageOrderNum:String;
		private var $totalCost:String;
		private var $singleItemQuantity:uint;
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
		
		public function get buyerPostCode():String 
		{
			return $buyerPostCode;
		}
		
		public function set buyerPostCode(value:String):void 
		{
			$buyerPostCode = value;
		}
		
		public function get weight():String 
		{
			return $weight;
		}
		
		public function set weight(value:String):void 
		{
			$weight = value;
		}
		
		public function get packageOrderNum():String 
		{
			return $packageOrderNum;
		}
		
		public function set packageOrderNum(value:String):void 
		{
			$packageOrderNum = value;
		}
		
		public function get totalCost():String 
		{
			return $totalCost;
		}
		
		public function set totalCost(value:String):void 
		{
			$totalCost = value;
		}
		
		public function get singleItemQuantity():uint 
		{
			return $singleItemQuantity;
		}
		
		public function set singleItemQuantity(value:uint):void 
		{
			$singleItemQuantity = value;
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