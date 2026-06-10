"""MS-OVBA compression algorithm (MS-OVBA 2.4.1.3.6).

VBA module sources in vbaProject.bin are stored compressed with this scheme.
oletools.olevba provides decompress_stream; this module adds compress_stream.

Used by extract_vba.py and build_vba.py.
"""
import struct
import math


def copytoken_help(decompressed_current, decompressed_chunk_start):
    difference = decompressed_current - decompressed_chunk_start
    if difference < 1:
        return 0x0FFF, 0xF000, 4, 4098
    bit_count = max(int(math.ceil(math.log(difference, 2))), 4)
    length_mask = 0xFFFF >> bit_count
    offset_mask = (~length_mask) & 0xFFFF
    maximum_length = (0xFFFF >> bit_count) + 3
    return length_mask, offset_mask, bit_count, maximum_length


def compress_chunk(chunk_data):
    """Compress a single chunk (up to 4096 bytes of input)."""
    n = len(chunk_data)
    body = bytearray()
    htbl = {}
    MAX_CANDIDATES = 32

    pos = 0
    while pos < n:
        flag_byte_pos = len(body)
        body.append(0)
        flag_byte = 0

        for bit_index in range(8):
            if pos >= n:
                break
            offset = 0
            length = 0
            if pos + 3 <= n:
                _, _, bit_count, max_length = copytoken_help(pos, 0)
                max_offset = 1 << bit_count
                key = (chunk_data[pos], chunk_data[pos + 1], chunk_data[pos + 2])
                cands = htbl.get(key, ())
                min_start = pos - max_offset
                best_off = 0
                best_len = 0
                for s in reversed(cands):
                    if s < min_start:
                        break
                    ml = 3
                    while ml < max_length and pos + ml < n and chunk_data[s + ml] == chunk_data[pos + ml]:
                        ml += 1
                    if ml > best_len:
                        best_len = ml
                        best_off = pos - s
                        if ml >= max_length:
                            break
                offset, length = best_off, best_len

            if offset > 0 and length >= 3:
                _, _, bit_count, _ = copytoken_help(pos, 0)
                copy_token = ((offset - 1) << (16 - bit_count)) | (length - 3)
                body.extend(struct.pack("<H", copy_token & 0xFFFF))
                flag_byte |= (1 << bit_index)
                for k in range(length):
                    p = pos + k
                    if p + 3 <= n:
                        kk = (chunk_data[p], chunk_data[p + 1], chunk_data[p + 2])
                        lst = htbl.setdefault(kk, [])
                        lst.append(p)
                        if len(lst) > MAX_CANDIDATES:
                            del lst[0:len(lst) - MAX_CANDIDATES]
                pos += length
            else:
                body.append(chunk_data[pos])
                if pos + 3 <= n:
                    kk = (chunk_data[pos], chunk_data[pos + 1], chunk_data[pos + 2])
                    lst = htbl.setdefault(kk, [])
                    lst.append(pos)
                    if len(lst) > MAX_CANDIDATES:
                        del lst[0:len(lst) - MAX_CANDIDATES]
                pos += 1

        body[flag_byte_pos] = flag_byte

    return bytes(body)


def compress_stream(uncompressed):
    """Compress bytes per MS-OVBA. Returns compressed bytes."""
    output = bytearray()
    output.append(0x01)  # signature byte
    pos = 0
    n = len(uncompressed)
    if n == 0:
        return bytes(output)
    while pos < n:
        chunk_end = min(pos + 4096, n)
        chunk_input = uncompressed[pos:chunk_end]
        body = compress_chunk(bytes(chunk_input))
        if len(body) < len(chunk_input):
            chunk_size_value = (2 + len(body)) - 3
            header = (chunk_size_value & 0x0FFF) | (0b011 << 12) | (1 << 15)
            output.extend(struct.pack("<H", header))
            output.extend(body)
        else:
            if len(chunk_input) == 4096:
                header = (4095 & 0x0FFF) | (0b011 << 12) | (0 << 15)
                output.extend(struct.pack("<H", header))
                output.extend(chunk_input)
            else:
                chunk_size_value = (2 + len(body)) - 3
                header = (chunk_size_value & 0x0FFF) | (0b011 << 12) | (1 << 15)
                output.extend(struct.pack("<H", header))
                output.extend(body)
        pos = chunk_end
    return bytes(output)
