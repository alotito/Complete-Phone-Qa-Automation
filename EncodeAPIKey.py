import base64
# Paste your freshly copied, original plaintext API key between the quotes below:
actual_api_key = "AIzaSyDC9H0Qth9brmlURsTiK77r2Bf3dEDdqfg"

# It's good to print the key you are about to encode to visually verify it one last time
# (but be careful not to leave this in a script that others might see if the key is sensitive)
print(f"DEBUG: About to encode this exact key: '{actual_api_key}'")
print(f"DEBUG: Length of key to encode: {len(actual_api_key)}")

encoded_bytes = base64.b64encode(actual_api_key.encode('utf-8'))
base64_encoded_key = encoded_bytes.decode('utf-8')

print(f"Base64 Encoded Key to put in config.ini: {base64_encoded_key}")
print(f"Encoded Length: {len(base64_encoded_key)}")