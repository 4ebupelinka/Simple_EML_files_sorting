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
    # List of characters that are not allowed in folder names
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
    # Replace invalid characters in the filename
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
        if filename.endswith(".eml"):  # Check if the file has .eml extension
            file_path = os.path.join(folder_path, filename)
            with open(file_path, "rb") as f:
                # Parse the .eml file
                msg = BytesParser(policy=policy.default).parse(f)
                sender = msg.get("From")  # Get sender from the message
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
    for sender in senders:
        # Sanitize sender name to create a folder
        sender_name = sanitize_folder_name(sender)
        sender_folder = os.path.join(output_path, sender_name)

        # Create folder if it does not exist
        if not os.path.exists(sender_folder):
            os.makedirs(sender_folder)

        # Copy emails from the sender to the corresponding folder
        for filename in os.listdir(folder_path):
            if filename.endswith(".eml"):  # Check if the file has .eml extension
                file_path = os.path.join(folder_path, filename)
                with open(file_path, "rb") as f:
                    # Parse the .eml file
                    msg = BytesParser(policy=policy.default).parse(f)
                    sender_in_msg = msg.get("From")  # Get sender from the message
                    subject = msg.get("Subject", "No Subject")  # Get subject (or "No Subject" if missing)
                    date_received = msg.get("Date")

                    # Format the received date
                    try:
                        date_received = email.utils.parsedate_to_datetime(date_received)
                        date_str = date_received.strftime("%Y-%m-%d_%H-%M-%S")
                    except Exception:
                        date_str = "UnknownDate"

                    # If the sender matches, copy the file to the appropriate folder
                    if sender_in_msg == sender:
                        # Sanitize subject and create a new filename
                        sanitized_subject = sanitize_filename(subject)
                        filename_to_copy = f"{sanitized_subject}_{date_str}.eml"

                        # Destination path for the new file
                        destination_path = os.path.join(sender_folder, filename_to_copy)

                        # Copy the file to the corresponding folder
                        shutil.copy(file_path, destination_path)
                        print(f"Copied email {filename} to folder {sender_name} with name {filename_to_copy}")


def main():
    # Path to the folder containing .eml files
    folder_path = "email"
    # Path to the folder where emails will be sorted
    output_path = "sort"

    # Get unique senders from all .eml files
    senders = get_senders_from_eml(folder_path)

    # Copy emails into folders for each sender
    copy_emails_to_sender_folders(folder_path, output_path, senders)


if __name__ == "__main__":
    main()
