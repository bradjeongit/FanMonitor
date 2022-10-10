
set_id = 3
set_cmd = [0xC9, 0x06, 0x00, 0x1C, 0x00, set_id]

print('\n' + 'id', str(set_id))


def crc16(data):
    data = bytearray(data)
    poly = 0xA001
    crc = 0xFFFF

    for b in data:
        crc ^= (0xFF & b)
        for _ in range(0, 8):
            if crc & 0x0001:
                crc = ((crc >> 1) & 0xFFFF) ^ poly
            else:
                crc = ((crc >> 1) & 0xFFFF)

    return crc


def calc_crc():
    set_code = ''

    for i in range(0, 6):
        temp = '00' + str(hex(set_cmd[i]))[2:]
        temp = temp.upper()
        set_code = set_code + temp[-2:]

    print('code', set_code)

    result = str(hex(crc16(set_cmd)))

#    print('result', result)
    crc = result[4:6] + result[2:4]

    print('crc ', crc, '\n')
    print('id set code', set_code + crc)

    print('--------------------------------------------------------------')


calc_crc()
