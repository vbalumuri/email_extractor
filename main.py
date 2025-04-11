import win32com.client  # pip install pywin32
import csv
import os

def main():
    # Specify the date range in MM/DD/YYYY HH:MM AM/PM format
    FROM_DATE = "01/01/2023 00:00 AM"  # inclusive start
    TO_DATE   = "02/01/2023 11:59 PM"  # inclusive end

    # Output CSV file path
    output_csv = r"C:\Path\To\unique_senders.csv"

    # Connect to Outlook MAPI
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox

    # Build a Restrict filter for items between FROM_DATE and TO_DATE
    # Ensure the format matches Outlook's expected "MM/DD/YYYY HH:MM[AM/PM]"
    filter_str = f"[ReceivedTime] >= '{FROM_DATE}' AND [ReceivedTime] <= '{TO_DATE}'"
    print(f"Applying Restrict filter: {filter_str}")
    filtered_items = inbox.Items.Restrict(filter_str)

    # Use a dictionary to track unique email addresses
    # Key = email address, Value = display name
    unique_senders = {}

    count = 0
    for msg in filtered_items:
        try:
            from_email = msg.SenderEmailAddress
            from_name = msg.SenderName
            if from_email:
                # Store/overwrite the display name for this address (if needed)
                unique_senders[from_email.lower()] = from_name
            count += 1
        except Exception as e:
            print(f"Error reading message: {e}")

    print(f"Scanned {count} messages in total.")
    print(f"Found {len(unique_senders)} unique sender addresses.")

    # Write the results to CSV (overwrite if already exists)
    os.makedirs(os.path.dirname(output_csv), exist_ok=True)
    with open(output_csv, mode="w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        # Write a header
        writer.writerow(["Email Address", "Display Name"])
        # Write each unique sender
        for email_addr, display_name in unique_senders.items():
            writer.writerow([email_addr, display_name or ""])

    print(f"Unique sender list written to: {output_csv}")

if __name__ == "__main__":
    main()