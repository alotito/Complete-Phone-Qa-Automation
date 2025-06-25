import hashlib
import os

def get_file_sha256(filepath):
    """
    Calculates the SHA256 hash of a file.

    Args:
        filepath (str): The path to the file.

    Returns:
        str: The hexadecimal SHA256 hash string, or None if an error occurs.
    """
    sha256_hash = hashlib.sha256()
    try:
        with open(filepath, "rb") as f:
            # Read and update hash string in chunks of 4K
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    except FileNotFoundError:
        print(f"Error: File not found at '{filepath}'")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

if __name__ == "__main__":
    file_to_hash = input("Enter the path to the file: ")

    # Basic validation for file existence, though the function also checks
    if not os.path.isfile(file_to_hash):
        print(f"Error: '{file_to_hash}' is not a valid file or does not exist.")
    else:
        print("\nCalculating SHA256 hash...")
        file_hash = get_file_sha256(file_to_hash)
        if file_hash:
            print(f"\nFile: {file_to_hash}")
            print(f"SHA256: {file_hash}")