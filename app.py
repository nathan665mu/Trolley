import os
from webapp.app import app  # re-export app for backward compatibility

if __name__ == "__main__":
    # Allow running `python app.py` from project root
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
