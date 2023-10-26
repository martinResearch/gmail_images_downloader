"""

Generate an Application-Specific Password:

Google recommends using an application-specific password for less secure apps or scripts. To generate one:

Go to your Google Account settings page: https://myaccount.google.com/.
IN the search bar at the top left, type "App passwords."
 sign in if prompted.
You may be asked to re-enter your Google password. Afterward, you can generate an application-specific password for your script.

Generate the Password:

In the "App passwords" section, select "Other (Custom name)" from the dropdown list.
Enter a name for your app or script (e.g., "Gmail Image Extractor").
Click "Generate."
You will receive a 16-character application-specific password. Keep this password secure.
"""

from datetime import datetime
import email
from email.header import decode_header

import imaplib
import io
import os
import re
import string

from tqdm import tqdm

email_address = "your_email@gmail.com"
password = "the app password you created in gmail"
folder = "c:/downloaded_images"
label = "inbox"


def decode_mime_encoded_word(encoded_word):
    decoded_words = decode_header(encoded_word)
    decoded_filename = ""
    for decoded, charset in decoded_words:
        if isinstance(decoded, bytes):
            decoded_filename += decoded.decode(charset if charset else "utf-8")
        elif isinstance(decoded, str):
            decoded_filename += decoded
    return decoded_filename


def clean_filename(filename):
    # Define a regular expression pattern for valid filename characters
    valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
    # Replace any character not in the valid set with an underscore
    cleaned_filename = "".join(c if c in valid_chars else "_" for c in filename)
    # Remove any repeated underscores
    cleaned_filename = re.sub("_+", "_", cleaned_filename)
    return cleaned_filename


def download_and_process_email(folder, email_id, mail, pbar):
    status = ""
    while not status == "OK":
        status, email_data = mail.fetch(email_id, "(RFC822)")

    # Parse the email data
    msg = email.message_from_bytes(email_data[0][1])

    # Extract the date of the email
    if isinstance(msg["Date"], email.header.Header):
        email_date_str = msg["Date"].__str__()
    else:
        email_date_str = msg["Date"]
    email_date_str = " ".join(email_date_str.split(" ")[:5])
    email_date = datetime.strptime(email_date_str, "%a, %d %b %Y %H:%M:%S")

    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get("Content-Disposition") is None:
                continue

            # Save the image attachments
            filename = part.get_filename()
            if filename is None:
                filename = f"noname.{part.get_content_subtype()}"

            filename = decode_mime_encoded_word(filename)
            filename = clean_filename(filename)
            if filename.endswith(".eml"):
                continue
            if filename:
                # Try to extract the date from the image metadata
                try:
                    from PIL import Image

                    img = Image.open(io.BytesIO(part.get_payload(decode=True)))
                    exif_data = img._getexif()
                    if 36867 in exif_data:
                        image_date = datetime.strptime(exif_data[36867], "%Y:%m:%d %H:%M:%S")
                    else:
                        image_date = email_date
                except Exception:
                    image_date = email_date

                # Create a prefix with the date
                date_prefix = image_date.strftime("%Y%m%d_%H%M%S")

                # Rename and save the image with the date prefix
                new_filename = f"{date_prefix}_{filename}"
                with open(os.path.join(folder, new_filename), "wb") as fp:
                    fp.write(part.get_payload(decode=True))
    pbar.update(1)


def download_images():
    # Your Gmail credentials

    # Connect to Gmail using IMAP
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(email_address, password)

    # Select the mailbox you want to fetch emails from
    mail.select(label)

    # Search for all emails with attachments
    status, email_ids = mail.search(None, 'X-GM-RAW "has:attachment filename:(jpg OR jpeg OR png OR gif)"')

    # Get the list of email IDs
    email_id_list = email_ids[0].split()

    print(f"found {len(email_id_list)} emails")

    # Create a directory to save the downloaded images
    os.makedirs(folder, exist_ok=True)

    # Create a tqdm progress bar
    with tqdm(total=len(email_id_list)) as pbar:
        for email_id in email_id_list:
            download_and_process_email(folder, email_id, mail, pbar)

    # Logout and close the connection
    mail.logout()


if __name__ == "__main__":
    download_images()
