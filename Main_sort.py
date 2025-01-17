import os
import shutil
import email
from email import policy
from email.parser import BytesParser
import re
from datetime import datetime

def sanitize_folder_name(name):
    """
    Sanitize folder name by replacing invalid characters with an underscore.

    Args:
        name (str): The folder name to sanitize.

    Returns:
        str: The sanitized folder name.
    """
    invalid_chars = r'<>:"/\\|?*'
    return re.sub(f"[{re.escape(invalid_chars)}]", "_", name)

def sanitize_filename(name):
    """
    Sanitize filename by replacing invalid characters with an underscore.

    Args:
        name (str): The filename to sanitize.

    Returns:
        str: The sanitized filename.
    """
    return re.sub(f"[{re.escape(r'<>:\"/\\|?*')}]|\n", "_", name)

def get_senders_from_eml(folder_path):
    """
    Get unique senders from all .eml files in the specified folder.

    Args:
        folder_path (str): The path to the folder containing .eml files.

    Returns:
        set: A set of unique senders.
    """
    senders = set()
    for filename in os.listdir(folder_path):
        if filename.endswith(".eml"):
            file_path = os.path.join(folder_path, filename)
            with open(file_path, "rb") as f:
                msg = BytesParser(policy=policy.default).parse(f)
                sender = msg.get("From")
                if sender:
                    senders.add(sender)
    return senders

def copy_emails_to_sender_folders(folder_path, output_path, senders):
    """
    Copy emails from each sender into separate folders based on sender email.

    Args:
        folder_path (str): The path to the folder containing .eml files.
        output_path (str): The path to the folder where sorted emails will be saved.
        senders (set): A set of unique senders.
    """
    total_files = sum(1 for filename in os.listdir(folder_path) if filename.endswith(".eml"))
    file_count = 0

    for sender in senders:
        sender_name = sanitize_folder_name(sender)
        sender_folder = os.path.join(output_path, sender_name)

        if not os.path.exists(sender_folder):
            os.makedirs(sender_folder)

        for filename in os.listdir(folder_path):
            if filename.endswith(".eml"):
                file_path = os.path.join(folder_path, filename)
                with open(file_path, "rb") as f:
                    msg = BytesParser(policy=policy.default).parse(f)
                    sender_in_msg = msg.get("From")
                    subject = msg.get("Subject", "No Subject")
                    date_received = msg.get("Date")

                    try:
                        date_received = email.utils.parsedate_to_datetime(date_received)
                        date_str = date_received.strftime("%Y-%m-%d_%H-%M-%S")
                    except Exception:
                        date_str = "UnknownDate"

                    if sender_in_msg == sender:
                        sanitized_subject = sanitize_filename(subject)
                        filename_to_copy = f"{sanitized_subject}_{date_str}.eml"
                        destination_path = os.path.join(sender_folder, filename_to_copy)

                        shutil.copy(file_path, destination_path)
                        file_count += 1
                        print(f"[{file_count}/{total_files}] Copied email {filename} to folder {sender_name} with name {filename_to_copy}")

def sort_emails():
    folder_path = "email"
    output_path = "sort"

    senders = get_senders_from_eml(folder_path)
    copy_emails_to_sender_folders(folder_path, output_path, senders)

if __name__ == "__main__":
    sort_emails()
