package rtfconverter

import (
	"fmt"
	"strings"
	"strconv"
)

type rtfElement interface {
   setParent(p *rtfGroup)
   GetParent() (*rtfGroup)
   Dump(level int)
}


/**
 * RTF Groups
 */
type rtfGroup struct {
	children []rtfElement
	parent *rtfGroup
}


func (r *rtfGroup) addChild(c rtfElement) {
	c.setParent(r)
	r.children = append(r.children, c)
}

func (r *rtfGroup) GetChildren() []rtfElement{
	return r.children
}

func (r *rtfGroup) setParent(p *rtfGroup) {
	r.parent = p
}

func (r *rtfGroup) GetParent()(*rtfGroup) {
	return r.parent
}


func (r *rtfGroup) IsDestination() bool {
	return r.CheckChildAtIndex(0, "*")
}

func (r *rtfGroup) IsRtfGroup() bool {
	return r.CheckChildAtIndex(0, "rtf")
}


/**
 * check if the group define the font table (first child must be fonttbl)
 */
func (r *rtfGroup) IsFontTable() bool {
	return r.CheckChildAtIndex(0, "fonttbl")
}

/**
 * check if the group define the stylesheet
 */
func (r *rtfGroup) IsStylesheet() bool {
	return r.CheckChildAtIndex(0, "stylesheet")
}

/**
 * check if the group define the listtables
 */
func (r *rtfGroup) IsListtables() bool {
	return r.CheckChildAtIndex(0, "listtables")
}

/**
 * check if info group - document info are
 */
func (r *rtfGroup) IsInfo() bool {
	return r.CheckChildAtIndex(0, "info")
}

/**
 * check if the group define the files table
 */
func (r *rtfGroup) IsFilesTable() bool {
	return (r.IsDestination() && r.CheckChildAtIndex(1, "filetbl"))
}

/**
 * check if the group define the revtbl
 */
func (r *rtfGroup) IsTrackChanges() bool {
	return (r.IsDestination() && r.CheckChildAtIndex(1, "revtbl")) || r.CheckChildAtIndex(0, "revtbl")
}



/**
 * check if the group define the font table (first child must be colortbl)
 */
func (r *rtfGroup) IsColorTable() bool {
	return r.CheckChildAtIndex(0, "colortbl")
}

/**
 * check if is a fontinfo entry from font table group
 * <fontnum><fontfamily><fcharset>?<fprq>?<panose>?<nontaggedname>?<fontemb>?<codepage>? <fontname><fontaltname>? ';'
 * eq:
 * 	{\f0\fswiss\fcharset0 Arial;}
 */
func (r *rtfGroup) IsFontInfo() bool {
	return r.CheckChildAtIndex(0, "f")
}


/**
 * check if the group is an alternative name
 * {\*\falt xxxx}
 */
func (r *rtfGroup) IsFontAlternative() bool {
	return r.IsDestination() && r.CheckChildAtIndex(1, "falt")
}



func (r *rtfGroup) CheckChildAtIndex(idx int, checkWord string)  (bool) {

	if idx < len(r.children) {
		child := r.children[idx]
	    // First child not a control symbol?
	    switch child.(type) {
	    	case *rtfControlSymbol:
	    		return child.(*rtfControlSymbol).symbol == checkWord
	    	case *rtfControlWord:
	    		return child.(*rtfControlWord).word == checkWord
	    }
	}
	return false;
}


func (r *rtfGroup) Dump(level int) {
	fmt.Printf("%sGroup (Children: %d)\r\n", strings.Repeat(" ", level), len(r.children));
	if len(r.children) > 0 {
		for _, child := range(r.children) {
			child.Dump(level+1);
		}
	}
}


/**
 * RTF Word Control
 * word property will saved the control word without \
 * parameter may contains negative numbers
 *
 * 	eg: \rtf1
 * 	 word: rtf
 * 	 parameter: 1
 *
 */
type rtfControlWord struct {
	word string
	parameter string
	parent *rtfGroup
}


func (r *rtfControlWord) setParent(p *rtfGroup) {
	r.parent = p
}

func (r *rtfControlWord) GetParent()(*rtfGroup) {
	return r.parent
}

func (r *rtfControlWord) GetWord()(string) {
	return r.word
}


/**
 * parameter should be an integer; I return it as a string so I can make the difference between 0 and empty
 * @param  {[type]} r *rtfControlWord) GetParameter() (string [description]
 * @return {[type]}   [description]
 */
func (r *rtfControlWord) GetParameter() (string) {
	return r.parameter
}

/**
 * for control words the default parameter is 1 and is integer
 */
func (r *rtfControlWord) GetIntParameter() (int) {
	if (r.parameter == "") {
		return 1
	}
	p,err := strconv.Atoi(r.parameter)
	if err == nil {
		return p
	}
	return 1
}

func (r *rtfControlWord) Dump(level int) {
	fmt.Printf("%sControl Word (Word: \\%s%v)\r\n", strings.Repeat(" ", level), r.word, r.parameter);
}


/**
 * symbols control
 * parameter is used only for \'HH control symbol where HH are hexadecimal digits
 * in the symbol property the value will be saved without the \
 * eg: \'HH =>
 * 		symbol: '
 * 		parameter: HH
 */
type rtfControlSymbol struct {
	symbol string
	parameter string
	parent *rtfGroup
}

func (r *rtfControlSymbol) GetSymbol()(string) {
	return r.symbol
}

func (r *rtfControlSymbol) setParent(p *rtfGroup) {
	r.parent = p
}

func (r *rtfControlSymbol) GetParent()(*rtfGroup) {
	return r.parent
}

/**
 * usually the parameter is empty, except when the control symbol is \'
 * @param  {[type]} r *rtfControlSymbol) GetParameter() (string [description]
 * @return {[type]}   [description]
 */
func (r *rtfControlSymbol) GetParameter() (string) {
	return r.parameter
}

func (r *rtfControlSymbol) Dump(level int) {
	fmt.Printf("%sControl Symbol (Symbol: \\%s%s)\r\n", strings.Repeat(" ", level), r.symbol, r.parameter);
}


/**
 * Text
 */
type rtfText struct {
	content []byte
	parent *rtfGroup
}


func (r *rtfText) setParent(p *rtfGroup) {
	r.parent = p
}

func (r *rtfText) GetParent()(*rtfGroup) {
	return r.parent
}

func (r *rtfText) GetContent()([]byte) {
	return r.content
}

func (r *rtfText) Dump(level int) {
	fmt.Printf("%sControl Text: %s)\r\n", strings.Repeat(" ", level), r.content);
}
