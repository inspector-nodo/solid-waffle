def detect_compression_magic_bytes(file_path):
    with open(file_path, "rb") as f:
        magic_bytes = f.read(4)  # Read the first 4 bytes
    
    magic_dict = {
        b'\x1F\x8B': "gzip",
        b'\x50\x4B\x03\x04': "zip",
        b'\x42\x5A\x68': "bzip2",
        b'\xFD\x37\x7A\x58': "xz",
        b'\x75\x73\x74\x61': "tar",
    }

    for magic, compression in magic_dict.items():
        if magic_bytes.startswith(magic):
            return compression
    
    return "Unknown"

file_path = "example.gz"
print(f"Compression Type: {detect_compression_magic_bytes(file_path)}")
