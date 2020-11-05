package rtfconverter

type rtfHtmlInterpreter struct {
	content []byte
}

func (p *rtfHtmlInterpreter) Parse(rtfObj RtfStructure) ([]byte, error) {

	var (
		result []byte
		err error
	)

	/**
	 *	The de-encapsulating RTF reader SHOULD<14> inspect no more than the first 10 RTF tokens
	 *	(that is, begin group marks and control words) in the input RTF document, in sequence, starting from the beginning of the RTF document.
	 *	If one of the control words is the FROMHTML control word, the de-encapsulating RTF reader SHOULD conclude that the RTF document contains
	 *	an encapsulated HTML document and stop further inspection. If one of the control words is the FROMTEXT control word, the de-encapsulating
	 *	RTF reader SHOULD conclude that the RTF document was produced from a plain text document and stop further inspection.
	 */
	if (rtfObj.IsHtmlEncapsulated()) {
		// the RTF was generated from a html file
		parser := rtfHtmlEncapsulatedInterpreter{styleTag: "span"}
		result, err =  parser.Parse(rtfObj)
	}

	return result, err
}

