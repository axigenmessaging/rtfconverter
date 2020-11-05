/**
 * decompose a rtf file into a structure of rtf tokens: rtf words, symbols, etc
 */

package rtfconverter


import (
	"io"
	"os"
	"bufio"
	"bytes"
	"strconv"
)

type RtfStructure struct {
	reader *bufio.Reader
	Root *rtfGroup
	currentGroup *rtfGroup

	/**
	 * This keyword represents the number of bytes corresponding to a given \uN Unicode character.
	 * This keyword may be used at any time, and values are scoped like character properties.
	 * That is, a \ucN keyword applies only to text following the keyword, and within the same (or deeper) nested braces.
	 * On exiting the group, the previous \uc value is restored. The reader must keep a stack of counts seen and use the most
	 * recent one to skip the appropriate number of characters when it encounters a \uN keyword. When leaving an RTF group which specified
	 * a \uc value, the reader must revert to the previous value. A default of 1 should be assumed if no \uc keyword has been seen in the current or outer scopes.
	 * A common practice is to emit no ANSI representation for Unicode characters within a Unicode destination context
	 * (that is, inside a \ud destination.).
	 * Typically, the destination will contain a \uc0 control sequence.
	 * There is no need to reset the count on leaving the \ud destination as the scoping rules will ensure the previous value is restored.
	 */

	uc []int
}

/**
 * load a file
 * @param {[type]} file string [description]
 */
func (rtfObj *RtfStructure) ParseFile(filename string) (error) {
	_, err := os.Stat(filename)
	if err != nil {
		return err
	}

	fileReader, err := os.Open(filename)
	defer fileReader.Close()

/*
	data, err := ioutil.ReadFile(filename)
	if err != nil {
		return err
	}
*/
	reader := bufio.NewReader(fileReader)

	rtfObj.setReader(reader)
	return rtfObj.Parse()
}

func (rtfObj *RtfStructure) ParseBytes(content []byte) (error) {

	btsReader := bytes.NewReader(content)

	reader := bufio.NewReader(btsReader)
	rtfObj.setReader(reader)
	return rtfObj.Parse()
}


func (rtfObj *RtfStructure) setReader(reader *bufio.Reader) {
	rtfObj.reader = reader
}


func (rtfObj *RtfStructure) Parse() (error){
	var (
		b byte
		err error
	)

	for {
		// read a single line

		b, err = rtfObj.reader.ReadByte()

		//fmt.Println("Read: ", string(b))

		if (rtfObj.currentGroup == nil && rtfObj.Root != nil) {
			// ignore text after RTF group tag is closed
			break
		}

		// What type of character is this?
		switch {
			case string(b) == "{":
			  rtfObj.startGroup();
			case string(b) == "}":
			  rtfObj.endGroup();
			case string(b) == "\\":
			  rtfObj.parseControl();
			default:
			  // move the pointer back 1 char
			  rtfObj.reader.UnreadByte()
			  rtfObj.parseText()
		}
		// If we're just at the EOF, break
		if err != nil || err == io.EOF {
		    break
		}

	}

	 return err

}

/**
 * create a new group if the current char is {, and add it as a child to current group, and set the current group the new created group
 *
 * @param  {[type]} rtfObj *RtfStructure)   parseStartGroup( [description]
 * @return {[type]}        [description]
 */
func (rtfObj *RtfStructure) startGroup() {
	//fmt.Println("start group")
	group := rtfGroup{}
	if rtfObj.Root == nil {
		rtfObj.Root = &group
		rtfObj.currentGroup = rtfObj.Root
		rtfObj.uc = append(rtfObj.uc, 1)
	} else {
		// add the new group as a child to the current one
		rtfObj.currentGroup.addChild(&group)

		// set the active group the new one
		rtfObj.currentGroup = &group

		// inherit the uc from the last group
		rtfObj.uc = append(rtfObj.uc, rtfObj.uc[len(rtfObj.uc)-1])
	}
}


/**
 * an end group is identified; set the parent of current group as current group
 */

func (rtfObj *RtfStructure) endGroup() {
	//fmt.Println("end group")

	// when a group is closed, set the new current group his parent
	rtfObj.currentGroup = rtfObj.currentGroup.GetParent()

	// uc for the group is lost when the group is closed, and the previous value is restored
	if (len(rtfObj.uc) > 0) {
		rtfObj.uc = rtfObj.uc[:len(rtfObj.uc)-1]
	}

}

func (rtfObj *RtfStructure) parseText() {

	var (
		b byte
		br rune
		bp []byte
		err error
	)
	buffer := &bytes.Buffer{}

	//fmt.Println("parse text")

	// continue read until meet a char that tell us the text ends (start group, end group, control word , control symbol)
	for {
		// peek to the next 2 bytes
		bp, err = rtfObj.reader.Peek(2)

		if err != nil && len(bp)==0 {
			// end of file
			break
		}

		//fmt.Printf("Peek Parse Text: \"%s\"\r\n", string(bp))

		// ignore EOL chars
		if (len(bp)>=1 && (string(bp[0]) == "\r" || string(bp[0]) == "\n" )) {
			_, err = rtfObj.reader.ReadByte()
			if err != nil {
				break
			}
			continue
		}


		if string(bp[0]) == "\\" {
			// check if the read char \ is an escape or mark a new control word
			if len(bp) == 2 && string(bp[1]) != "}" && string(bp[1]) != "{" && string(bp[1]) != "\\" {
				// it's the beginning of a new control word
				break
			} else {
				/**
				 * the char "\" escape a char; we have to read both to avoid considering the escaped char as a stop token char
				 */

				// read the escape char (\)
				b, err = rtfObj.reader.ReadByte()
				if err != nil {
					break
				}
				buffer.WriteByte(b)

				// read the escaped char
				b, err = rtfObj.reader.ReadByte()
				if err != nil {
					break
				}
				buffer.WriteByte(b)

				continue
			}
		} else if string(bp[0]) == "{" || string(bp[0])=="}" {
			// it's end of a group
			break
		}

		// if it get here it means is a text char
		br, _, err = rtfObj.reader.ReadRune()
		if err != nil {
			break;
		}
		buffer.WriteRune(br)
	}

	if buffer.Len() > 0 {
		// append text token only if the text if not empty
		obj := &rtfText{content: buffer.Bytes()}

		//fmt.Println("write text: ", string(obj.content))
		rtfObj.currentGroup.addChild(obj)
	}

}


/**
 * the char before the reader pointer is \ that tell us is a control symbol (a single non alphanumeric char) or control word
 * @param  {[type]} rtfObj *RtfStructure)   parseControl( [description]
 * @return {[type]}        [description]
 */
func (rtfObj *RtfStructure) parseControl() {

	//fmt.Println("Parse control")

	isControlWord := true
	// continue read until meet a char that tell us the text ends (start group, end group, control word , control symbol)

	// peek to the next 2 bytes
	bp, err := rtfObj.reader.Peek(1)

	if (err != nil) {
		// the peek is probably end of file
		return
	}
	if !ByteIsAsciiLetter(bp[0]) {
		// it's not a alphanumeric char => it's a control symbol
		isControlWord = false
    }

	if isControlWord {
		//fmt.Println("parse control word")
		rtfObj.parseControlWord()
	} else {
		//fmt.Println("parse control symbol")
		rtfObj.parseControlSymbol()
	}
}

/**
 * control symbols tokens that begin with a non alphanumeric char, formed by a single char and without parameters
 * if the char is ', that means it will follow a 2 digit hex number.
 *
 * The function match:
 * \NonAlphaNumericChar (\~, \*, \)
 * \'HH , HH - hexadecimal value
 *
 * if the function detects \\n or \\r it will consider it's a new line escaped and tranform it to controlWord \par
 *
 * @param  {[type]} rtfOj *RtfStructure)   parseControlSymbol( [description]
 * @return {[type]}       [description]
 */
func (rtfObj *RtfStructure) parseControlSymbol() {
	var bp []byte

	  parameterBuffer := &bytes.Buffer{}


      b, err := rtfObj.reader.ReadByte()

      if (err != nil) {
      	return
      }

      // Symbols ordinarily have no parameter. However,
      // if this is \', then it is followed by a 2-digit hex-code:
      // Treat EOL symbols as \par control word

      if (string(b) == "\r" || string(b) == "\n") {
      	// check if the next char is \r or \n, so we can remove the entire \r\n pair (it's not a symbol, it's an escaped end of line??)
		bp, err := rtfObj.reader.Peek(1)

		if err != nil {
			// probably is end of file
			return
		}

		if (len(bp)==1 &&  (string(bp[0]) == "\r" || string(bp[0]) == "\n")) {
			_, err := rtfObj.reader.ReadByte()
			if err != nil {
				// probably is end of file
				return
			}
		}

        obj := &rtfControlWord{word: "par"}
        rtfObj.currentGroup.addChild(obj)
        return;

      } else if(string(b) == "'") {
      	// detected \'HH
      	// read next 2 chars (must be 2 digits, reprezenting a hex number)
        bp, err = rtfObj.reader.Peek(1)
        if (err == nil && ByteIsHexDigit(bp[0])) {
        	// read first hexDigit
    		rtfObj.reader.ReadByte()
    		parameterBuffer.WriteByte(bp[0])

    		// read second hexDigit
    		bp, err = rtfObj.reader.Peek(1)
    		if (err == nil && ByteIsHexDigit(bp[0])) {
	    		rtfObj.reader.ReadByte()
	    		parameterBuffer.WriteByte(bp[0])
    		}
        }
      }

      objSymbol := &rtfControlSymbol{symbol: string(b), parameter: parameterBuffer.String()};
      rtfObj.currentGroup.addChild(objSymbol)
}


/**
 * control words patterns are:
 * 		\letters[numericParameter]
 * 		\u[-]NNNNN
 *
 * @param  {[type]} rtfOj *RtfStructure)   parseControlWord( [description]
 * @return {[type]}       [description]
 */
func (rtfObj *RtfStructure) parseControlWord() {

		var (
			b byte
			bp []byte
			br rune
			err error
			parameter int
		)

		wordBuffer := &bytes.Buffer{}
		parameterBuffer := &bytes.Buffer{}

		// extract the word
		for {
			bp, err = rtfObj.reader.Peek(1)
			if (err != nil) {
				break
			}

			if (ByteIsAsciiLetter(bp[0])) {
				b,err = rtfObj.reader.ReadByte()
				//fmt.Println("Read extract word: ",string(b))
				if (err != nil) {
					break
				}
				wordBuffer.WriteByte(b)
			} else {
				break
			}
		}

		// check if the parameter is negative (-digits)
		bp, err = rtfObj.reader.Peek(1)
		if (err == nil && string(bp[0]) == "-") {
			b, err = rtfObj.reader.ReadByte()
			//fmt.Println("Read extract negative: ",string(b))
			if err == nil {
				parameterBuffer.WriteString("-")
			}
		}

		// extract the numeric parameter
		for {
			bp, err = rtfObj.reader.Peek(1)
			if err != nil || !ByteIsDigit(bp[0]) {
				break;
			}

			b, err = rtfObj.reader.ReadByte()

			//fmt.Println("Read extract parameter: ",string(b))

			if err != nil {
				break
			}
			parameterBuffer.WriteByte(b)
		}

		bp, err = rtfObj.reader.Peek(1)
		if (err == nil && string(bp[0]) == " ") {
			/**
			 * if a space delimits the control word, the space does not appear in the document.
			 * Any characters following the delimiter, including spaces, will appear in the document.
			 * For this reason, you should use spaces only where necessary; do not use spaces merely to break up RTF code.
			 *
			 *  depite the above explanation , I remove the space from documente and control word
			 */

			rtfObj.reader.ReadByte()
		}


		controlWord := wordBuffer.String()
		controlWordParameter := parameterBuffer.String()


		if len(controlWordParameter) > 0 {
			parameter, err = strconv.Atoi(controlWordParameter)
			if (err  != nil) {
				parameter = 1
			}
		} else {
			parameter = 1
		}

		/**
		 * This keyword represents the number of bytes corresponding to a given \uN Unicode character.
		 * This keyword may be used at any time, and values are scoped like character properties.
		 * That is, a \ucN keyword applies only to text following the keyword, and within the same (or deeper) nested braces.
		 * On exiting the group, the previous \uc value is restored. The reader must keep a stack of counts seen and use the most
		 * recent one to skip the appropriate number of characters when it encounters a \uN keyword. When leaving an RTF group which specified
		 * a \uc value, the reader must revert to the previous value. A default of 1 should be assumed if no \uc keyword has been seen in the current or outer scopes.
		 * A common practice is to emit no ANSI representation for Unicode characters within a Unicode destination context
		 * (that is, inside a \ud destination.).
		 * Typically, the destination will contain a \uc0 control sequence.
		 * There is no need to reset the count on leaving the \ud destination as the scoping rules will ensure the previous value is restored.
		 */
        if (controlWord == "uc") {
        	// rewrite the last entry from uc - by default when a new group is added the entry is automatically added with value 1
            rtfObj.uc[len(rtfObj.uc)-1] = parameter;
        }

		/*
		* This keyword represents a single Unicode character which has no equivalent ANSI representation based on the current ANSI code page. N represents the Unicode character value expressed as a decimal number.
		* This keyword is followed immediately by equivalent character(s) in ANSI representation. In this way, old readers will ignore the \uN keyword and pick up the ANSI representation properly. When this keyword is encountered, the reader should ignore the next N characters, where N corresponds to the last \ucN value encountered.
		* As with all RTF keywords, a keyword-terminating space may be present (before the ANSI characters) which is not counted in the characters to skip. While this is not likely to occur (or recommended), a \bin keyword, its argument, and the binary data that follows are considered one character for skipping purposes. If an RTF scope delimiter character (that is, an opening or closing brace) is encountered while scanning skippable data, the skippable data is considered to be ended before the delimiter. This makes it possible for a reader to perform some rudimentary error recovery. To include an RTF delimiter in skippable data, it must be represented using the appropriate control symbol (that is, escaped with a backslash,) as in plain text. Any RTF control word or symbol is considered a single character for the purposes of counting skippable characters.
		* An RTF writer, when it encounters a Unicode character with no corresponding ANSI character, should output \uN followed by the best ANSI representation it can manage. Also, if the Unicode character translates into an ANSI character stream with count of bytes differing from the current Unicode Character Byte Count, it should emit the \ucN keyword prior to the \uN keyword to notify the reader of the change.
		* RTF control words generally accept signed 16-bit numbers as arguments. For this reason, Unicode values greater than 32767 must be expressed as negative numbers.
		*/

        if (controlWord == "u") {
            // Convert parameter to unsigned decimal unicode
            if (parameter < 0) {
                parameter += 65536;
            }

            // Will ignore replacement characters uc times
            uc := rtfObj.uc[len(rtfObj.uc)-1];
            for {

                bp ,err = rtfObj.reader.Peek(1)
                if err != nil {
                	break
                }
 				if (string(bp[0]) == "{" || string(bp[0]) == "}") {
                	// start / ends a group
                    break;
                }

                // read an entire char (if is multibyte)
                br,_, err = rtfObj.reader.ReadRune()
                if err != nil {
                	break
                }

                // peek to the next char
                bp ,err = rtfObj.reader.Peek(1)
                if err != nil {
                	break
                }
                // If the replacement character is encoded as hexadecimal value \'HH then jump over it
                if (string(br) == "\\" && string(bp[0]) == "'") {
                    // move pointer after 4 chars \'HH
                    _, err = rtfObj.reader.ReadByte()
                    if err != nil {
                    	break
                    }
                    _, err = rtfObj.reader.ReadByte()
                    if err != nil {
                    	break
                    }
                    _, err = rtfObj.reader.ReadByte()
                    if err != nil {
                    	break
                    }
                }

                uc--;

                if uc <= 0 {
                	break
                }
            }
        }

		obj := &rtfControlWord{word: controlWord, parameter: controlWordParameter}

		rtfObj.currentGroup.addChild(obj)
}

func (rtfObj *RtfStructure) Dump() {
	rtfObj.Root.Dump(0)
}


/**
 * the root group must have the first control word == rtf1
 * @param  {[type]} rtfObj *RtfStructure) IsValid() (bool [description]
 * @return {[type]}        [description]
 */
func (rtfObj *RtfStructure) IsValid() (bool) {
	children := rtfObj.Root.GetChildren()
	if len(children) > 0 {
		switch children[0].(type) {
		case *rtfControlWord:
			if children[0].(*rtfControlWord).GetWord() == "rtf" && children[0].(*rtfControlWord).GetParameter()=="1" {
				return true
			}
		}
	}
	return false
}

/**
 *	The de-encapsulating RTF reader SHOULD<14> inspect no more than the first 10 RTF tokens
 *	(that is, begin group marks and control words) in the input RTF document, in sequence, starting from the beginning of the RTF document.
 *	If one of the control words is the FROMHTML control word, the de-encapsulating RTF reader SHOULD conclude that the RTF document contains
 *	an encapsulated HTML document and stop further inspection. If one of the control words is the FROMTEXT control word, the de-encapsulating
 *	RTF reader SHOULD conclude that the RTF document was produced from a plain text document and stop further inspection.
 */

func (rtfObj *RtfStructure) IsHtmlEncapsulated() (bool) {
	idx := 0;
	children := rtfObj.Root.GetChildren()
	for {
		for _, item := range children {
			switch item.(type) {
			case *rtfControlWord:
				if item.(*rtfControlWord).GetWord() == "fromhtml" {
					return true
				}
			}
			idx++
		}

		if (idx > 10) {
			break;
		}
	}
	return false
}


func (rtfObj *RtfStructure) IsTextEncapsulated() (bool) {
	idx := 0;
	children := rtfObj.Root.GetChildren()
	for {
		for _, item := range children {
			switch item.(type) {
			case *rtfControlWord:
				if item.(*rtfControlWord).GetWord() == "fromtext" {
					return true
				}
			}
			idx++
		}

		if (idx > 10) {
			break;
		}
	}
	return false
}