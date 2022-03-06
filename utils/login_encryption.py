from Crypto.Cipher import AES
from binascii import b2a_hex, a2b_hex


def login_encrypt(message, key):
    mode = AES.MODE_OFB
    cryptor = AES.new(key.encode('utf-8'), mode, b'0000000000000000')
    length = 16
    count = len(message)
    if count % length != 0:
        add = length - (count % length)
    else:
        add = 0
    message = message + ('\0' * add)
    ciphertext = cryptor.encrypt(message.encode('utf-8'))
    result = b2a_hex(ciphertext)
    return result.decode('utf-8')


def login_decrypt(result, key):
    mode = AES.MODE_OFB
    cryptor = AES.new(key.encode('utf-8'), mode, b'0000000000000000')
    plain_text = cryptor.decrypt(a2b_hex(result))
    return plain_text.decode('utf-8').rstrip('\0')

# en = login_encrypt("我爱python", key='KosatlRPAclientpirvateky')
# print(en)
# de = login_decrypt("918d035e6c9dab0ee824d0fe4ce7635f992ec292", key='KosatlRPAclientpirvateky')
# print(de)
