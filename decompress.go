package rtfconverter

import (
	"errors"
)

var CRC32_TABLE []int

func init() {
	CRC32_TABLE = make([]int, 256)
	for i := 0; i < 256; i += 1 {
		c := i
		for j := 0; j < 8; j += 1 {
			if (c & 1) == 1 {
				c = 0xEDB88320 ^ shiftZeroLeft(c, 1)
			} else {
				c = shiftZeroLeft(c, 1)
			}
		}
		CRC32_TABLE[i] = c
	}
}

func Decompress(src []byte) ([]byte, error) {
	const (
		MAGIC_COMPRESSED   = 0x75465a4c
		MAGIC_UNCOMPRESSED = 0x414c454d
		DICT_SIZE          = 4096
		DICT_MASK          = DICT_SIZE - 1 // for quick modulo operations
	)

	var (
		dst []byte     // destination for uncompressed bytes
		in  int    = 0 // current position in src array
		out int    = 0 // current position in dst array
	)

	prebuf :=
		"{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}" +
			"{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript " +
			"\\fdecor MS Sans SerifSymbolArialTimes New RomanCourier" +
			"{\\colortbl\\red0\\green0\\blue0\n\r\\par " +
			"\\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx"

	COMPRESSED_RTF_PREBUF := []byte(prebuf)

	// Get header fields
	if src == nil || len(src) < 16 {
		return nil, errors.New("Invalid compressed-RTF header")
	}

	compressedSize := int(getU32(src, in))
	in += 4

	uncompressedSize := int(getU32(src, in))
	in += 4

	magic := int(getU32(src, in))
	in += 4

	// Note: CRC must be validated only for compressed data (and includes padding)
	crc32sum := int(getU32(src, in))
	in += 4

	if compressedSize != len(src)-4 {
		// Check size excluding the size field itself
		return nil, errors.New("compressed data size mismatch")
	}

	// Process the data
	if magic == MAGIC_UNCOMPRESSED {
		return src, nil
	} else if magic == MAGIC_COMPRESSED {
		if crc32sum != calculateCRC32(src, 16, len(src)-16) {
			return nil, errors.New("compressed-RTF CRC32 failed")
		}

		out = len(COMPRESSED_RTF_PREBUF)
		dst = make([]byte, out+uncompressedSize)

		for i, v := range COMPRESSED_RTF_PREBUF {
			dst[i] = v
		}

		var flagCount int = 0
		var flags int = 0

		for {

			// Each flag byte controls 8 literals/references, 1 per bit
			// Each bit is 1 for reference, 0 for literal
			if (flagCount & 7) == 0 {
				flags = int(src[in])
				in += 1
			} else {
				flags = flags >> 1
			}
			flagCount += 1

			if (flags & 1) == 0 {
				dst[out] = src[in]
				out += 1
				in += 1

			} else {

				// Read reference: 12-bit offset (from block start) and 4-bit length
				offset := int(src[in]) & 0xFF
				in += 1

				length := int(src[in]) & 0xFF
				in += 1

				offset = (offset << 4) | shiftZeroLeft(length, 4) // the offset from block start
				length = (length & 0xF) + 2                       // the number of bytes to copy

				// The decompression buffer is supposed to wrap around back
				// to the beginning when the end is reached. we save the
				// need for such a buffer by pointing straight into the data
				// buffer, and simulating this behaviour by modifying the
				// pointers appropriately.
				offset = out & ^DICT_MASK | offset // the absolute offset in array

				if offset >= out {
					if offset == out {
						break // a self-reference marks the end of data
					}
					offset -= DICT_SIZE // take from previous block
				}
				// Note: can't use System.arraycopy, because the referenced
				// bytes can cross through the current out position.
				end := offset + length

				for offset < end {
					dst[out] = dst[offset]
					out += 1
					offset += 1
				}
			}
		}

		return dst[len(COMPRESSED_RTF_PREBUF):], nil

	} else { // unknown magic number
		return nil, errors.New("Unknown compression type (magic number " + string(magic) + ")")
	}

	return dst, nil
}

/**
 * Returns an unsigned 32-bit value from little-endian ordered bytes.
 *
 * @param buf a byte array from which byte values are taken
 * @param offset the offset within buf from which byte values are taken
 * @return an unsigned 32-bit value as a long
 */
func getU32(buf []byte, offset int) int64 {
	t0 := int64(buf[offset])
	t1 := int64(buf[offset+1])
	t2 := int64(buf[offset+2])
	t3 := int64(buf[offset+3])

	b := (t0 & 0xFF) | (t1&0xFF)<<8 | (t2&0xFF)<<16 | (t3&0xFF)<<24

	return b & int64(0xFFFFFFFF)
}

/**
 * Calculates the CRC32 of the given bytes, continuing a previous calculation.
 * <p>
 * The CRC32 calculation is similar to the standard one as demonstrated
 * in RFC 1952, but with the inversion (before and after the calculation)
 * omitted.
 *
 * @param buf the byte array to calculate CRC32 on
 * @param off the offset of buf at which the CRC32 calculation will start
 * @param len the number of bytes on which to calculate the CRC32
 * @param crc the previous CRC32 value calculated from preceding bytes
 * @return the CRC32 value
 */
func calculateCRC32Continue(buf []byte, off int, l int, crc int) int {
	end := off + l
	for i := off; i < end; i += 1 {
		crc = CRC32_TABLE[(crc^int(buf[i]))&0xFF] ^ shiftZeroLeft(crc, 8)
	}
	return crc
}

/**
 * Calculates the CRC32 of the given bytes.
 * <p>
 * The CRC32 calculation is similar to the standard one as demonstrated
 * in RFC 1952, but with the inversion (before and after the calculation)
 * omitted.
 *
 * @param buf the byte array to calculate CRC32 on
 * @param off the offset of buf at which the CRC32 calculation will start
 * @param len the number of bytes on which to calculate the CRC32
 * @return the CRC32 value
 */
func calculateCRC32(buf []byte, off int, l int) int {
	return calculateCRC32Continue(buf, off, l, 0)
}

// The >>> Java equivalent
func shiftZeroLeft(i int, positions uint) int {
	return i >> positions
}
