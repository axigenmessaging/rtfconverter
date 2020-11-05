package rtfconverter

import (
	"regexp"
	"errors"
    "golang.org/x/text/encoding"
    "golang.org/x/text/encoding/charmap"
    "golang.org/x/text/encoding/japanese"
    "golang.org/x/text/encoding/simplifiedchinese"
    "golang.org/x/text/encoding/korean"
    "golang.org/x/text/encoding/traditionalchinese"
)

// check if a byte is lower letter to check if it is part of a control word
func ByteIsAsciiLetter(b byte) bool {
	isLetter, _ := regexp.MatchString(`[A-Za-z]{1}`, string(b));
	return isLetter
}

func ByteIsDigit(b byte) bool {
	isDigit, _ := regexp.MatchString(`[0-9]{1}`, string(b));
	return isDigit
}

func ByteIsHexDigit(b byte) bool {
	isDigit, _ := regexp.MatchString(`[A-Za-z0-9]{1}`, string(b));
	return isDigit
}


var rtfEncodeCodePageMap map[string]string = map[string]string{
  "ansi" : "CP1252",
  "mac"  : "MAC",
  "pc"   : "CP437",
  "pca"  : "CP850",
  "437" : "CP437", // United States IBM
  "708" : "ASMO-708", // also [ISO-8859-6][ARABIC] Arabic
  /*  Not supported by iconv
  709, : "" // Arabic (ASMO 449+, BCON V4)
  710, : "" // Arabic (transparent Arabic)
  711, : "" // Arabic (Nafitha Enhanced)
  720, : "" // Arabic (transparent ASMO)
  */
  "819" : "CP819",   // Windows 3.1 (US and Western Europe)
  "850" : "CP850",   // IBM multilingual
  "852" : "CP852",   // Eastern European
  "860" : "CP860",   // Portuguese
  "862" : "CP862",   // Hebrew
  "863" : "CP863",   // French Canadian
  "864" : "CP864",   // Arabic
  "865" : "CP865",   // Norwegian
  "866" : "CP866",   // Soviet Union
  "874" : "CP874",   // Thai
  "932" : "CP932",   // Japanese
  "936" : "CP936",   // Simplified Chinese
  "949" : "CP949",   // Korean
  "950" : "CP950",   // Traditional Chinese
  "1250" : "CP1250",  // Windows 3.1 (Eastern European)
  "1251" : "CP1251",  // Windows 3.1 (Cyrillic)
  "1252" : "CP1252",  // Western European
  "1253" : "CP1253",  // Greek
  "1254" : "CP1254",  // Turkish
  "1255" : "CP1255",  // Hebrew
  "1256" : "CP1256",  // Arabic
  "1257" : "CP1257",  // Baltic
  "1258" : "CP1258",  // Vietnamese
  "1361" : "CP1361",   // Johab
}

func GetEncodingFromCodepage(code string) (string, error) {
	if val, ok := rtfEncodeCodePageMap[code]; ok {
	    return val, nil
	}

	return "", errors.New("Encoding Code Page not found")
}


var rtfEncodingCharsetMap map[int]string = map[int]string {
  0   : "CP1252", // ANSI: Western Europe
  1   : "CP1252", //*Default
  2   : "CP1252", //*Symbol
  3   : "",     // Invalid
  77  : "MAC",    //*also [MacRoman]: Macintosh
  128 : "CP932",  //*or [Shift_JIS]?: Japanese
  129 : "CP949",  //*also [UHC]: Korean (Hangul)
  130 : "CP1361", //*also [JOHAB]: Korean (Johab)
  134 : "CP936",  //*or [GB2312]?: Simplified Chinese
  136 : "CP950",  //*or [BIG5]?: Traditional Chinese
  161 : "CP1253", // Greek
  162 : "CP1254", // Turkish (latin 5)
  163 : "CP1258", // Vietnamese
  177 : "CP1255", // Hebrew
  178 : "CP1256", // Simplified Arabic
  179 : "CP1256", //*Traditional Arabic
  180 : "CP1256", //*Arabic User
  181 : "CP1255", //*Hebrew User
  186 : "CP1257", // Baltic
  204 : "CP1251", // Russian (Cyrillic)
  222 : "CP874",  // Thai
  238 : "CP1250", // Eastern European (latin 2)
  254 : "CP437",  //*also [IBM437][437]: PC437
  255 : "CP437", //*OEM still PC437
}

func GetEncodingFromCharset(code int) (string, error) {
	if val, ok := rtfEncodingCharsetMap[code]; ok {
	    return val, nil
	}

	return "", errors.New("Encoding Code Page not found")
}

var rtfHighlightMap map[int]string = map[int]string {
    1 : "Black",
    2 : "Blue",
    3 : "Cyan",
    4 : "Green",
    5 : "Magenta",
    6 : "Red",
    7 : "Yellow",
    8 : "Unused",
    9 :  "DarkBlue",
    10 : "DarkCyan",
    11 : "DarkGreen",
    12 : "DarkMagenta",
    13 : "DarkRed",
    14 : "DarkYellow",
    15 : "DarkGray",
    16 : "LightGray",
}


func ConvertToUtf8(b []byte, srcEncoding string) ([]byte, error) {
  var (
    result []byte
    err error
    dec  *encoding.Decoder
  )

  //fmt.Println("Decoding: ", string(b))

  switch (srcEncoding) {
    case "MAC": // [MacRoman]: Macintosh
        dec = charmap.Macintosh.NewDecoder()
    case "CP437": // United States IBM
        dec = charmap.CodePage437.NewDecoder()
    case "ASMO-708": // also [ISO-8859-6][ARABIC] Arabic
        dec = charmap.ISO8859_6.NewDecoder()
    case "CP819":   // Windows 3.1 (US and Western Europe)
        dec = charmap.ISO8859_1.NewDecoder()
    case "CP850":   // IBM multilingual
        dec = charmap.CodePage850.NewDecoder()
    case "CP852":   // Eastern European
        dec = charmap.CodePage852.NewDecoder()
    case "CP860":   // Portuguese
        dec = charmap.CodePage860.NewDecoder()
    case "CP862":   // Hebrew
        dec = charmap.CodePage862.NewDecoder()
    case "CP863":   // French Canadian
        dec = charmap.CodePage863.NewDecoder()
    case "CP864":   // Arabic
    case "CP865":   // Norwegian
        dec = charmap.CodePage865.NewDecoder()
    case "CP866":   // Soviet Union
        dec = charmap.CodePage866.NewDecoder()
    case "CP874":   // Thai
        dec = charmap.Windows874.NewDecoder()
    case "CP932":   // Japanese
        dec = japanese.ShiftJIS.NewDecoder()
    case "CP936":   // Simplified Chinese
        dec = simplifiedchinese.GBK.NewDecoder()
    case "CP949":   // Korean
        dec = korean.EUCKR.NewDecoder();
    case "CP950":   // Traditional Chinese
        dec = traditionalchinese.Big5.NewDecoder();
    case "CP1250":  // Windows 3.1 (Eastern European)
        dec =  charmap.Windows1250.NewDecoder()
    case "CP1251":  // Windows 3.1 (Cyrillic)
        dec =  charmap.Windows1251.NewDecoder()
    case "CP1252":  // Western European
        dec =  charmap.Windows1252.NewDecoder()
    case "CP1253":  // Greek
        dec =  charmap.Windows1253.NewDecoder()
    case "CP1254":  // Turkish
        dec =  charmap.Windows1254.NewDecoder()
    case "CP1255":  // Hebrew
        dec =  charmap.Windows1255.NewDecoder()
    case "CP1256":  // Arabic
        dec =  charmap.Windows1256.NewDecoder()
    case "CP1257":  // Baltic
        dec =  charmap.Windows1257.NewDecoder()
    case "CP1258":  // Vietnamese
        dec =  charmap.Windows1258.NewDecoder()
    case "CP1361":   // Johab
        dec = korean.EUCKR.NewDecoder()

  }

  if dec != nil {
    result, err = dec.Bytes(b)
  } else {
    result = b[0:]
  }

  return result, err
}