package application;

import java.util.ArrayList;
import java.util.List;

public class Util {
	
	public int decodeToDecimal(String columnStr){
		columnStr.toUpperCase();
		char[] eachChar = new char[columnStr.length()];
		columnStr.getChars(0, columnStr.length(), eachChar, 0);
		
		int columnNo = 0;
		for(int i=0; i < eachChar.length; i++){
			int charInt = eachChar[i];
			charInt -=65;
			columnNo += charInt*Math.pow(26, i);
		}

		return columnNo;
	}
	
	public String encodeToAlphaberic(int columnNo){
		List<Character> alps = new ArrayList<Character>();
		// �ʿ��ϰ� �Ǹ� �����Ѵ�
		
		
		
		String columnStr = null;
		return columnStr;
	}
	
	
}
