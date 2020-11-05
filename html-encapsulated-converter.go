/*
	https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxrtfex/906fbb0f-2467-490e-8c3e-bdc31c5e9d35
*/

package rtfconverter

import (
	"errors"
	"bytes"
	"strconv"
	"fmt"
	"encoding/binary"
)

var rtfFontsHtmlMap map[string]string = map[string]string {
	"fnil": "", 	//Unknown or default fonts (the default)
	"froman": "serif",	//Roman, proportionally spaced serif fonts	Times New Roman, Palatino
	"fswiss": "sans-serif",	//Swiss, proportionally spaced sans serif fonts	Arial
	"fmodern": "monospace", //	Fixed-pitch serif and sans serif fonts	Courier New, Pica
	"fscript": "cursive", //	Script fonts	Cursive
	"fdecor": "fantasy",	// Decorative fonts	Old English, ITC Zapf Chancery
	"ftech": "",	// Technical, symbol, and mathematical fonts	Symbol
	"fbidi": "",	//Arabic, Hebrew, or o
}


type rtfHtmlEncapsulatedInterpreter struct {
	content 				bytes.Buffer
	insideHtmlTagGroup 		int
	rtfEncoding 			string
	defaultFont 			int
	fontTable  				map[int]*rtfFontTableItem
	colorTable 				[]rtfColor
	styleTag		        string

	// keep the initial state when a group is parsed
	groupsInitialStates				[]rtfState

	// keep the previous state of the parsed group
	groupPreviousState		rtfState

	// keep the current state of the parsed group
	groupCurrentState		rtfState

	// count how many style tag are opened
	styleTagOpened int

	bodyStarted bool
	bodyStopped bool
}


type rtfFontTableItem struct {
    charsetIndex int
    familyCode string
    familyName string
    familyAlternativeName string
}

type rtfState struct {
	states map[string]string
}

func NewRtfState() (rtfState) {
	c := rtfState{
		states: map[string]string{
			//"f": "0", // default font
			//"fs": "24", // default font size
		},
	}

	return c
}

/**
 * make a copy of the structure
 * @param  {[type]} c rtfState)     copy( [description]
 * @return {[type]}   [description]
 */
func (c rtfState) copy() (rtfState){
	c1 := NewRtfState()

	for i,v := range c1.states {
		c1.states[i] = v
	}

	return c1
}

func (c rtfState) stateExists(word string) (bool){
	if _, ok := c.states[word]; ok {
		return true
	}

	return false
}

func (c rtfState) stateValue(word string) (string){
	if v, ok := c.states[word]; ok {
		return v
	}

	return ""
}

/**
 * check if 2 states are the same for styling
 * @param  {[type]} c rtfState)     hasSameStyle(c1 rtfState) (rtfState [description]
 * @return {[type]}   [description]
 */
func (c rtfState) hasSameStyle(c1 rtfState) (bool) {
	cStyleStates := 0
	c1StyleStates := 0

	for stateWord,stateValue := range c.states {
		if stateType, _ := DetectWordState(stateWord); stateType == "style" {
			cStyleStates++
			if rStateValue, ok := c1.states[stateWord]; ok {
				if rStateValue != stateValue {
					return false
				}
			} else {
				// the c1 do not contains this tag
				return false
			}
		}
	}

	for rStateWord,_ := range c.states {
		if stateType, _ := DetectWordState(rStateWord); stateType == "style" {
			c1StyleStates++
		}
	}

	if c1StyleStates != cStyleStates {
		return false
	}


	return true
}

type rtfColor struct {
	r int
	g int
	b int
}

/**
 * return the html hex code of the color
 * @param  {[type]} c *rtfColor)    getHexCode( [description]
 * @return {[type]}   [description]
 */
func (c *rtfColor) getHexCode() (string){
	return fmt.Sprintf("#%x%x%x", c.r, c.g, c.b)
}


func (p *rtfHtmlEncapsulatedInterpreter) Parse(rtfObj RtfStructure) ([]byte, error) {

	p.content = bytes.Buffer{}
	p.insideHtmlTagGroup = 0

	if (!rtfObj.IsValid()) {
		return nil, errors.New("The RTF file is not valid.")
	}

	p.parseElement(rtfObj.Root)

	return p.content.Bytes(), nil
}

/**
 * detect the rtf element from structure and decide the parser
 * @param  {[type]} p *rtfHtmlEncapsulatedInterpreter) parseElement(item rtfElement [description]
 * @return {[type]}   [description]
 */
func (p *rtfHtmlEncapsulatedInterpreter) parseElement(item rtfElement) {
	switch item.(type) {
		case *rtfGroup:
			p.parseGroup(item.(*rtfGroup));
		case *rtfControlSymbol:
			p.parseControlSymbol(item.(*rtfControlSymbol));
		case *rtfControlWord:
			p.parseControlWord(item.(*rtfControlWord));
		case *rtfText:
			p.parseText(item.(*rtfText));

	}
}

/**
 * parse a rtf group
 * @param  {[type]} p *rtfHtmlEncapsulatedInterpreter) parseGroup(item *rtfGroup [description]
 * @return {[type]}   [description]
 */
func (p *rtfHtmlEncapsulatedInterpreter) parseGroup(item *rtfGroup) {
	children := item.GetChildren()

	isHtmlTagDestinationGroup := false;
	isStartBodyTag := false

	// check if we parse a destination group that is a htmltag (a group where the first 2 childs are \*\htmltag)
	if (item.IsDestination()) {
		if (len(children)>=2) {
			switch children[1].(type) {
			case *rtfControlWord:
				if (children[1].(*rtfControlWord).GetWord() == "htmltag") {
					isHtmlTagDestinationGroup = true
					if !p.bodyStarted && children[1].(*rtfControlWord).GetIntParameter() == 50 {
						isStartBodyTag = true
					}
					if children[1].(*rtfControlWord).GetIntParameter() == 58 {
						p.bodyStopped = true
					}
				}
			}
		}
	}

	if isHtmlTagDestinationGroup {
		p.insideHtmlTagGroup++
	}

	if item.IsFontTable() {
		p.parseFontTableGroup(item)
	} else if item.IsColorTable() {
		p.parseColorTableGroup(item)
	} else if (item.IsStylesheet() || item.IsTrackChanges() || item.IsInfo() || item.IsListtables() || item.IsFilesTable()) {
		// ignore all these groups
	} else {

		// if the first group
		if (len(p.groupsInitialStates) == 0) {
			p.groupCurrentState = NewRtfState()
		}

		// save the state of the previous group
		p.groupsInitialStates = append(p.groupsInitialStates, p.groupCurrentState.copy())


		// when an state is open, try to close the previous one if
		//p.closeState()

		// everytime when a group start, we open a state that will be close when exit group
		//p.openState()

		for _, child := range children {
			p.parseElement(child)
		}

		// when a group end, we closed the state opened at the beginning, and restore previous group state
		//p.closeState()

		// restore current group to previous group states
		p.groupCurrentState = p.groupsInitialStates[len(p.groupsInitialStates)-1]

		// remove the last state from intial states array, so the last element will be the initial state of the current group
		p.groupsInitialStates = p.groupsInitialStates[:len(p.groupsInitialStates)-1]

		if (isStartBodyTag) {
			// we can add the styles
			p.bodyStarted = true
		}

	}

	if isHtmlTagDestinationGroup {
		p.insideHtmlTagGroup--
		if p.insideHtmlTagGroup < 0  {
			p.insideHtmlTagGroup = 0;
		}
	}
}


func (p *rtfHtmlEncapsulatedInterpreter) openState() {

	if !p.bodyStarted || p.bodyStopped {
		return
	}

	style := bytes.Buffer{}

	style.WriteString("style=\"")

	for stateWord, stateValue := range p.groupCurrentState.states {
		switch (stateWord) {
			case "b":
				if (stateValue == "1") {
					style.WriteString("text-weight: bold;");
				}
			case "i":
				if (stateValue == "1") {
					style.WriteString("text-weight: italic;");
				}
			case "v":
				if (stateValue == "0") {
					style.WriteString("display:none;");
				}
			case "f":
				fontIdx, err := strconv.Atoi(stateValue)

				if err == nil {
					if fItem, ok := p.fontTable[fontIdx]; ok {
		    			style.WriteString("font-family:")
		    			style.WriteString(fItem.familyName)
		    			style.WriteString(";")
	    			}
				}
			case "fs":
				if stateValue !="" && stateValue != "0" {
					style.WriteString("font-size:")
					style.WriteString(stateValue)
					style.WriteString("px;")
				}
			case "ul":
				if stateValue == "1" {
					style.WriteString("text-decoration:underline;");
				}
			case "strike":
				if stateValue == "1" {
					style.WriteString("text-decoration:line-through;");
				}
			case "cf","chcfpat":
				colorIdx, err := strconv.Atoi(stateValue)
				if err == nil && len(p.colorTable) >= colorIdx+1 {
					style.WriteString("color:");
					style.WriteString(p.colorTable[colorIdx].getHexCode());
					style.WriteString(";");
				}
			case "cb", "chcbpat", "highlight":
				colorIdx, err := strconv.Atoi(stateValue)
				if err == nil && len(p.colorTable) >= colorIdx+1 {
					style.WriteString("background-color:");
					style.WriteString(p.colorTable[colorIdx].getHexCode());
					style.WriteString(";");
				}
			case "super":
				if stateValue == "1" {
					style.WriteString("vertical-align: super;");
				}


		}
	}

	style.WriteString("\"")


	/**
	 *  save the last group state (the state configuration at the end group) to be used as current group initial state
	 *  init the current group state and group previous state with the initial group state
	 */

	p.openTag(p.styleTag, style.String())
}

func (p *rtfHtmlEncapsulatedInterpreter) closeState() {
	if !p.bodyStarted || p.bodyStopped {
		return
	}

	p.styleTagOpened++
	p.closeTag(p.styleTag)
}


func (p *rtfHtmlEncapsulatedInterpreter) openTag(tag string, attr string) {
	p.content.WriteString("<")
	p.content.WriteString(tag)
	p.content.WriteString(" ")
	p.content.WriteString(attr)
	p.content.WriteString(">")
}

func (p *rtfHtmlEncapsulatedInterpreter) closeTag(tag string) {
	if p.styleTagOpened > 0 {
		p.content.WriteString("</")
		p.content.WriteString(tag)
		p.content.WriteString(">")
		p.styleTagOpened--
	}
}

/**
 * 	extract font table
 *   {' \fonttbl (<fontinfo> | ('{' <fontinfo> '}'))+ '}'
 *   <fontnum><fontfamily><fcharset>?<fprq>?<panose>?
 *   <nontaggedname>?<fontemb>?<codepage>? <fontname><fontaltname>? ';'
 *
 * @param  {[type]} p *rtfHtmlEncapsulatedInterpreter) parseFontTableGroup(item *rtfGroup [description]
 * @return {[type]}   [description]
 */
func (p *rtfHtmlEncapsulatedInterpreter) parseFontTableGroup(item *rtfGroup) {
	p.fontTable = map[int]*rtfFontTableItem{}
	for _, child := range item.children {
		switch child.(type) {
		case *rtfGroup:
			if child.(*rtfGroup).IsFontInfo() {
				p.parseFontInfoGroup(child.(*rtfGroup))
			}
		}
	}
}
func (p *rtfHtmlEncapsulatedInterpreter) parseFontInfoGroup(item *rtfGroup) {
	var (
		fontIdx int
	)

	for _, child := range item.GetChildren() {
		switch cobj := child.(type) {
			case *rtfControlWord:
				switch cobj.GetWord() {
					case "f":
						fontIdx = cobj.GetIntParameter()
						p.fontTable[fontIdx] = &rtfFontTableItem{}
					case "fnil", "froman", "fswiss", "fmodern", "fscript", "fdecor", "ftech", "fbidi":
						// font fammily
						if ftItem, ok := p.fontTable[fontIdx]; ok {
							ftItem.familyCode = 	cobj.GetWord()
						}
					case "fcharset":
						if ftItem, ok := p.fontTable[fontIdx]; ok {
							ftItem.charsetIndex = cobj.GetIntParameter()
						}
				}
			case *rtfText:
				if ftItem, ok := p.fontTable[fontIdx]; ok {
					ftItem.familyName = string(bytes.TrimRight(cobj.GetContent(), ";"))
				}
			case *rtfGroup:
				if (cobj.IsFontAlternative()) {
					// check alternative font
					if len(cobj.children)>=3 {
						switch cobj.children[2].(type) {
							case *rtfText:
								if ftItem, ok := p.fontTable[fontIdx]; ok {
									ftItem.familyAlternativeName = string(cobj.children[2].(*rtfText).GetContent())
								}
						}

					}
				}
		}
	}
}

/**
 * extract colors from colortbl tag
 *  {\colortbl;\red0\green0\blue0;}
 * Index 0 of the RTF color table  is the 'auto' color
 * @param  {[type]} p *rtfHtmlEncapsulatedInterpreter) parseColorTableGroup(item *rtfGroup [description]
 * @return {[type]}   [description]
 */
func (p *rtfHtmlEncapsulatedInterpreter) parseColorTableGroup(item *rtfGroup) {
	if (!item.IsColorTable()) {
		return
	}

	color := rtfColor{
		r: 0,
		g: 0,
		b: 0,
	}
	for _, child := range item.GetChildren() {
		switch child.(type) {
			case *rtfControlWord:
				switch (child.(*rtfControlWord).GetWord()) {
					case "red":
						color.r = child.(*rtfControlWord).GetIntParameter()
					case "green":
						color.g = child.(*rtfControlWord).GetIntParameter()
					case "blue":
						color.b = child.(*rtfControlWord).GetIntParameter()
				}

			case *rtfText:
				// an end of color if marked by a ; text
				p.colorTable = append(p.colorTable, color)

				// reset color
				color = rtfColor{
					r: 0,
					g: 0,
					b: 0,
				}
		}
	}
}

func (p *rtfHtmlEncapsulatedInterpreter) parseControlSymbol(item *rtfControlSymbol) {

	if (p.groupCurrentState.stateExists("htmlrtf") && p.groupCurrentState.stateValue("htmlrtf") == "1") {
		/* Outside of an HTMLTAG destination groupIgnore and skip any text and RTF control words that are suppressed
		by any HTMLRTF control word other than the \fN control word. The de-encapsulating RTF reader SHOULD track the
		current font even when the corresponding \fN control word is inside of a fragment that is disabled with an HTMLRTF control word.
		*/
		return
	}

	switch item.GetSymbol() {
		case "'":
			// convert the string, reprezenting an hex number to a decimal number
			v, err := strconv.ParseInt(item.GetParameter(), 16, 16)

			if (err == nil) {
				b := make([]byte, 2)
				binary.LittleEndian.PutUint16(b, uint16(v))
				b = bytes.TrimRight(b, "\x00")
				r, _ := ConvertToUtf8(b, p.rtfEncoding);
				p.content.Write(r)
			}

			/*
			p.content.WriteString("%x")
			p.content.WriteString(item.GetParameter())
			*/
	}

	if (p.insideHtmlTagGroup > 0) {
		switch item.GetSymbol() {
			case "~":
				p.content.WriteString("&nbsp;")
			case "_":
				p.content.WriteString("&shy;")
		}
	}
}

/**
 * update the current state of a group
 * when the current state is updated, the previousState is closed, and a new state is opened
 *
 */

func (p *rtfHtmlEncapsulatedInterpreter) updateState(word string, value string) {
	stateScenario, stateValueType  := DetectWordState(word)

	var stateValue string

	switch (stateValueType) {
		case "flag":
			if (value == "0") {
				stateValue = "0" // off
			} else {
				stateValue = "1" // on
			}
			break;
		case "value":
			stateValue = value
	}

	// update the current state with the new value
	p.groupCurrentState.states[word] = stateValue

	if stateScenario == "style" {
		// close previous state
		//p.closeState();

		// open the new state
		//p.openState()
	}
}

func (p *rtfHtmlEncapsulatedInterpreter) parseControlWord(item *rtfControlWord) {
	switch item.GetWord() {
		case  "htmlrtf":
			p.updateState(item.GetWord(), item.GetParameter())
			return
	}

	if (p.insideHtmlTagGroup > 0) {
		// no control words will be added if are inside an htmltag
		switch  item.GetWord() {
			case "u" :
				p.content.WriteString("&#")
				p.content.WriteString(item.GetParameter())
				p.content.WriteString(";")
				return
			case "lquote":
				p.content.WriteString("&lsquo;")
				return
			case "rquote":
				p.content.WriteString("&rsquo;")
				return
			case "ldblquote":
				p.content.WriteString("&ldquo;")
				return
			case "rdblquote":
				p.content.WriteString("&rdquo;")
				return
			case "bullet":
				p.content.WriteString("&bull;")
				return
			case "endash":
				p.content.WriteString("&ndash;")
				return
			case "emdash":
				p.content.WriteString("&mdash;")
				return
		}
	} else {

		if (item.word != "f" && p.groupCurrentState.stateExists("htmlrtf") && p.groupCurrentState.stateValue("htmlrtf") == "1") {
			/* Outside of an HTMLTAG destination groupIgnore and skip any text and RTF control words that are suppressed
			by any HTMLRTF control word other than the \fN control word. The de-encapsulating RTF reader SHOULD track the
			current font even when the corresponding \fN control word is inside of a fragment that is disabled with an HTMLRTF control word.
			*/
			return
		}

		// outside html group
		switch  item.GetWord() {
			case "line" : // new line
				//p.content.WriteString("<br>")
				return
			case "par" :
				p.content.WriteString("\r\n")
				return
			case "tab" : // tab
				p.content.WriteString("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
				return
			case "deff":
	        	p.defaultFont = item.GetIntParameter()
	        case "ansi","mac","pc","pca":
	        	p.rtfEncoding, _ = GetEncodingFromCodepage(item.GetWord())
	        	return
	        case "ansicpg":
	        	 if item.GetIntParameter()>0 {
	        		p.rtfEncoding, _ = GetEncodingFromCodepage(item.GetParameter())
	        	}
	        	return
	        case "f", "fs":
	        	// font
	        	p.updateState(item.GetWord(), item.GetParameter())
	        	return
	     }
	}
}

func (p *rtfHtmlEncapsulatedInterpreter) parseText(item *rtfText) {
	if (p.groupCurrentState.stateExists("htmlrtf") && p.groupCurrentState.stateValue("htmlrtf") == "1") {
		/* Outside of an HTMLTAG destination groupIgnore and skip any text and RTF control words that are suppressed
		by any HTMLRTF control word other than the \fN control word. The de-encapsulating RTF reader SHOULD track the
		current font even when the corresponding \fN control word is inside of a fragment that is disabled with an HTMLRTF control word.
		*/
		return
	}

	// ignore any text outside an htmlTag group
	t, _ := ConvertToUtf8(item.GetContent(), p.rtfEncoding)
	p.content.Write(t)
}


/**
 * some rtfControlWord are states; this function return an scope and a value type for a control word
 * @param stateScope string - there are states used to modify style or other (like htmlrtf) => return scope
 * @param stateValueType value of a state can be with a on/off  value (flag) or a value (value)
 *
 * @TODO:
 * 	- the words added are only for decoding an rtf that was encoded from a html source
 */

func DetectWordState(word string) (stateScope string, stateValueType string) {
	switch (word) {
		case "htmlrtf": // HTMLRTF control word identifies fragments of RTF that were not in the original HTML content
			stateScope = "other"
			stateValueType = "flag"
		case "b", 	// bold
			"i",	// italic
			"v",	// hidden
			"ul",  // underline start
			"uldone",  // underline stop
			"strike",  // strike
			"super": // sup
			stateScope = "style"
			stateValueType = "flag"
		case "f", // font index from font table
			 "fs", // Font size in half-points (the default is 24)

			 "chcfpat", // N is the color of the background pattern, specified as an index into the document's color table (Character Borders and Shading) - we can use it as font color
			 "cf", // foreground color (the default is 0) - font color

			 "cb", //Background color (the default is 0)
			 "chcbpat": // N is the fill color, specified as an index into the document's color table. -  (Character Borders and Shading) - we can use it as backgorund color
			stateScope = "style"
			stateValueType = "value"
	}
	return
}