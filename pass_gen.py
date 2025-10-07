from streamlit_authenticator.utilities.hasher import Hasher

# Just pass a single string
hashed_password = Hasher.hash("macharlabhanu169")

print(hashed_password)

