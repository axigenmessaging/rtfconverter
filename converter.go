/**
 * converts RTF to text or html based on the RTF source
 **/

package rtfconverter

import (
	"errors"
	"io/ioutil"
//	"fmt"
)

type RtfInterpreter interface {
	Parse(rtfObj RtfStructure) ([]byte, error)
}


type rtfConverter struct {
	rtfObj RtfStructure
}



/**
 * create a new convertor
 */

func NewConverter() (rtfConverter) {
	c := rtfConverter{}
	return c;
}


func (c *rtfConverter) LoadFile(sourceFile string) {
	c.rtfObj = RtfStructure{}

	// decompose RTF into structure based on words, symbols, etc
	c.rtfObj.ParseFile(sourceFile)
}

func (c *rtfConverter) SetBytes(content []byte)  {
	c.rtfObj = RtfStructure{}

	// decompose RTF into structure based on words, symbols, etc
	c.rtfObj.ParseBytes(content)
}

func (c *rtfConverter) SaveFile(content []byte, path string) (error) {
	err := ioutil.WriteFile(path, content , 0644)

	return err
}


func (c *rtfConverter) Convert(exportType string) (result []byte, err error) {
	var (
		parser RtfInterpreter
	)

	parser, err = c.getInterpreter(exportType)

	if err != nil {
		return result, err
	}

	result, err = parser.Parse(c.rtfObj)

	if err != nil {
		return nil, err
	}

	return result, nil
}

func (c *rtfConverter) getInterpreter(interpreterType string) (RtfInterpreter, error) {
	switch interpreterType {
	case "html":
		return &rtfHtmlInterpreter{}, nil
	case "text":
		return &rtfTextInterpreter{}, nil
	default:
		return nil, errors.New("Parser for conversion do not exists.")
	}
}
