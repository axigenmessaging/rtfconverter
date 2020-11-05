/*
	https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxrtfex/906fbb0f-2467-490e-8c3e-bdc31c5e9d35
*/

package rtfconverter

import (
	"errors"
	"bytes"
	"strconv"
	"encoding/binary"
)

type rtfTextEncapsulatedInterpreter struct {
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


func (p *rtfTextEncapsulatedInterpreter) Parse(rtfObj RtfStructure) ([]byte, error) {

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
 * @param  {[type]} p *rtfTextEncapsulatedInterpreter) parseElement(item rtfElement [description]
 * @return {[type]}   [description]
 */
func (p *rtfTextEncapsulatedInterpreter) parseElement(item rtfElement) {
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
 * @param  {[type]} p *rtfTextEncapsulatedInterpreter) parseGroup(item *rtfGroup [description]
 * @return {[type]}   [description]
 */
func (p *rtfTextEncapsulatedInterpreter) parseGroup(item *rtfGroup) {
	children := item.GetChildren()

	if item.IsFontTable() {
		p.parseFontTableGroup(item)
	} else if item.IsColorTable() {
		p.parseColorTableGroup(item)
	} else if (item.IsStylesheet() || item.IsTrackChanges() || item.IsInfo() || item.IsListtables() || item.IsFilesTable()) {
		// ignore all these groups
	} else {
		if !item.IsDestination() {
			for _, child := range children {
				p.parseElement(child)
			}
		}
	}

}


/**
 * 	extract font table
 *   {' \fonttbl (<fontinfo> | ('{' <fontinfo> '}'))+ '}'
 *   <fontnum><fontfamily><fcharset>?<fprq>?<panose>?
 *   <nontaggedname>?<fontemb>?<codepage>? <fontname><fontaltname>? ';'
 *
 * @param  {[type]} p *rtfTextEncapsulatedInterpreter) parseFontTableGroup(item *rtfGroup [description]
 * @return {[type]}   [description]
 */
func (p *rtfTextEncapsulatedInterpreter) parseFontTableGroup(item *rtfGroup) {
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
func (p *rtfTextEncapsulatedInterpreter) parseFontInfoGroup(item *rtfGroup) {
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
 * @param  {[type]} p *rtfTextEncapsulatedInterpreter) parseColorTableGroup(item *rtfGroup [description]
 * @return {[type]}   [description]
 */
func (p *rtfTextEncapsulatedInterpreter) parseColorTableGroup(item *rtfGroup) {
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

func (p *rtfTextEncapsulatedInterpreter) parseControlSymbol(item *rtfControlSymbol) {
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
		case "~":
			p.content.WriteString("-")
		case "_":
			p.content.WriteString("_")
	}
}


func (p *rtfTextEncapsulatedInterpreter) parseControlWord(item *rtfControlWord) {
	// no control words will be added if are inside an htmltag
	switch  item.GetWord() {
		case "u" :
			i, err := strconv.Atoi(item.GetParameter())
			if err == nil {
				buff := new(bytes.Buffer)
         		binary.Write(buff, binary.LittleEndian, i)
				r, _ := ConvertToUtf8(buff.Bytes(), p.rtfEncoding);
				p.content.Write(r);
				//p.content.WriteRune(rune(i))
			}
			return
		case "lquote":
			p.content.WriteString("'")
			return
		case "rquote":
			p.content.WriteString("'")
			return
		case "ldblquote":
			p.content.WriteString("\"")
			return
		case "rdblquote":
			p.content.WriteString("\"")
			return
		case "bullet":
			//p.content.WriteString("&bull;")
			return
		case "endash":
			p.content.WriteString("-")
			return
		case "emdash":
			p.content.WriteString("--")
			return

		case "line" : // new line
			p.content.WriteString("\r\n")
			return
		case "par" :
			p.content.WriteString("\r\n")
			return
		case "tab" : // tab
			p.content.WriteString("\t")
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
        	return
     }
}

func (p *rtfTextEncapsulatedInterpreter) parseText(item *rtfText) {
	// ignore any text outside an htmlTag group
	t, _ := ConvertToUtf8(item.GetContent(), p.rtfEncoding)
	p.content.Write(t)
}
