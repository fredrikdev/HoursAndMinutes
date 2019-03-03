public class CustomAlgorithm {
	public static void main(String[] args) {
		// First parameter is the user's full name
		// Second parameter is the user's company name
		// Third parameters is the user's email address
		// Fourth parameter is the register date (any format)
		if (args.length < 4){
			System.out.println("ERROR");
		} else {
			// Code to produce registration information
			final String strMasterMask = "9EAZD1K70X24YR52V3C32HN4JI7G8051FBMLO5P61QS7T02U556WQW49FAS423MC";
			final String strUserMask =   "0f129874lhaf25dqwerAsdASFPHWE12463049dsajxzpoq5213rvsb23gweqwet2";
			final String strHMVer = "1.X.X";
			String strFullName = args[0];
			String strCompanyName = args[1];
			String strEMail = args[2];
			String strDate = args[3];
			String strUser = strFullName.trim() + " " + strCompanyName.trim() + " " + strEMail.trim() + " " + strHMVer.trim() + " " + strDate;
			String strRegCode = "";
			int y, x;

			// make sure that we got a long user name
			x = 0;
			y = (int)strUser.charAt(0);
			while (strUser.length() < 80) {
				strUser += strUserMask.charAt(y % strUserMask.length());
				y = (int)strUser.charAt(y % strUser.length());
				x += 3;
				y += x;
			}
			strUser = strUser.substring(0, 72);

			// create a long string of data
			strRegCode = "-- BEGIN LICENSE --\r\n" + "Hours and Minutes " + strHMVer + " " + strDate + "\r\n" + strFullName + "\r\n" + strCompanyName + "\r\n" + strEMail + "\r\n";
			strRegCode += doHash(strUser, strMasterMask);
			strRegCode +="\r\n-- END LICENSE   --";

			System.out.println(strRegCode);
		}
	}

	public static String doHash(String strString, String strMask) {
		String strResult01 = "";
		String strResult02 = "";
		String strResult03 = "";
		String strResult = "";
		int x;
		int y;
		int z;
		byte bytChar, bytHashChar;

		// pass 1 (hash jump encoding)
		for (y = 0; y < strString.length(); y++) {
			// get position to move to
			x = (byte)strString.charAt(y);
			x = x % strString.length();

			// get byte from this position
			bytChar = (byte)strString.charAt(x);

			// get hash char to use
			strResult01 += "" + strMask.charAt(Math.abs(bytChar - strString.length()) % strMask.length());
		}

		// pass 2 (lame encoding)
		for (y = 0; y < strString.length(); y++) {
			// get byte from this position
			bytChar = (byte)strString.charAt(y);

			// get hash char to use
			strResult02 += "" + strMask.charAt(bytChar % strMask.length());
		}

		// pass 3 (combine strings)
		x = 0;
		y = 0;
		while ((x < strResult01.length()) || (y < strResult02.length())) {
			if (x < strResult01.length()) strResult03 += strResult01.charAt(x);
			if (y < strResult02.length()) strResult03 += strResult02.charAt(y);
			x++;
			y++;
		}

		// pass 4 (even split)
		z = (int)(strResult03.length()/3);
		for (y = 0; y < strResult03.length(); y++) {
			if (((y % z) == 0) && (y != 0)) strResult += "\r\n";
			strResult += strResult03.charAt(y);
		}

		return strResult;
	}
}

