import customtkinter as ctk
from exchangelib import Credentials, Account, Configuration, DELEGATE
import threading
import queue
import time

class ExchangeService:
    def __init__(self):
        self.account = None

    def connect(self, server, username, password, impersonated_user=None):
        try:
            credentials = Credentials(username, password)
            config = Configuration(server=server, credentials=credentials)

            if impersonated_user:
                self.account = Account(
                    primary_smtp_address=impersonated_user,
                    config=config,
                    access_type=DELEGATE,
                    autodiscover=False
                )
            else:
                self.account = Account(
                    primary_smtp_address=username,
                    config=config,
                    access_type=DELEGATE,
                    autodiscover=False
                )
            # Test connection by getting the root folder
            self.account.root.get()
            return "Connection successful!"
        except Exception as e:
            return f"Error: {e}"

    def get_folders(self):
        if not self.account:
            return []

        folders_data = []
        try:
            # Using walk() to recursively get all folders
            for folder in self.account.root.walk():
                folders_data.append({
                    "name": folder.name,
                    "total_count": folder.total_count,
                    "obj": folder
                })
        except Exception as e:
            print(f"Error getting folders: {e}") # This should probably go to a log queue
        return folders_data

    def empty_folder(self, folder_obj, progress_callback):
        if not self.account:
            return

        total_items = folder_obj.total_count
        if total_items == 0:
            if progress_callback:
                progress_callback(0, 0, 0, 0, "Folder is already empty.")
            return

        processed_count = 0
        start_time = time.time()

        if progress_callback:
            progress_callback(0, total_items, 0, 0, "Starting to empty folder...")

        while processed_count < total_items:
            try:
                # Fetch items in batches
                items_to_delete = list(folder_obj.all().order_by('datetime_received')[:100]) # batch size of 100

                if not items_to_delete:
                    break # No more items to delete

                # Delete items one by one to provide granular progress
                for item in items_to_delete:
                    try:
                        item.delete()
                        processed_count += 1

                        if progress_callback:
                            elapsed_time = time.time() - start_time
                            items_per_second = processed_count / elapsed_time if elapsed_time > 0 else 0
                            remaining_items = total_items - processed_count
                            estimated_time_remaining = remaining_items / items_per_second if items_per_second > 0 else 0

                            progress_callback(
                                processed_count,
                                remaining_items,
                                items_per_second,
                                estimated_time_remaining,
                                f"Deleted item {processed_count}/{total_items}"
                            )
                    except Exception as e:
                        if progress_callback:
                            progress_callback(processed_count, total_items - processed_count, 0, 0, f"Error deleting an item: {e}")

            except Exception as e:
                if progress_callback:
                    progress_callback(processed_count, total_items - processed_count, 0, 0, f"Error fetching items: {e}")
                time.sleep(2) # Wait before retrying

        if progress_callback:
            progress_callback(processed_count, 0, 0, 0, "Folder cleaning complete.")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("EWS Folder Cleaner")
        self.geometry("1000x700")

        self.exchange_service = ExchangeService()
        self.log_queue = queue.Queue()
        self.worker_thread = None
        self.selected_folder_obj = None

        self.create_widgets()
        self.after(100, self.process_log_queue)

    def create_widgets(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Left Frame for folders
        self.folder_frame = ctk.CTkScrollableFrame(self, label_text="Folders")
        self.folder_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ns")

        # Right Frame for controls and logs
        self.right_frame = ctk.CTkFrame(self)
        self.right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        self.right_frame.grid_columnconfigure(0, weight=1)
        self.right_frame.grid_rowconfigure(2, weight=1)

        # Login widgets
        self.login_frame = ctk.CTkFrame(self.right_frame)
        self.login_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.login_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.login_frame, text="Server:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.server_entry = ctk.CTkEntry(self.login_frame, placeholder_text="outlook.office365.com")
        self.server_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(self.login_frame, text="Username:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.username_entry = ctk.CTkEntry(self.login_frame, placeholder_text="user@example.com")
        self.username_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(self.login_frame, text="Password:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.password_entry = ctk.CTkEntry(self.login_frame, show="*")
        self.password_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(self.login_frame, text="Impersonate:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.impersonate_entry = ctk.CTkEntry(self.login_frame, placeholder_text="(optional)")
        self.impersonate_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        self.connect_button = ctk.CTkButton(self.login_frame, text="Connect", command=self.connect)
        self.connect_button.grid(row=4, column=0, columnspan=2, pady=10)

        # Controls
        self.controls_frame = ctk.CTkFrame(self.right_frame)
        # self.controls_frame is not displayed initially

        self.start_button = ctk.CTkButton(self.controls_frame, text="Start Cleaning", command=self.start_cleaning, state="disabled")
        self.start_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(self.controls_frame)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        self.status_label = ctk.CTkLabel(self.controls_frame, text="Status: Disconnected")
        self.status_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")

        # Log Textbox
        self.log_textbox = ctk.CTkTextbox(self.right_frame, state="disabled")
        self.log_textbox.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")

    def log(self, message):
        self.log_queue.put(message)

    def process_log_queue(self):
        while not self.log_queue.empty():
            message = self.log_queue.get_nowait()
            self.log_textbox.configure(state="normal")
            self.log_textbox.insert("end", message + "\n")
            self.log_textbox.see("end")
            self.log_textbox.configure(state="disabled")
        self.after(100, self.process_log_queue)

    def connect(self):
        server = self.server_entry.get()
        username = self.username_entry.get()
        password = self.password_entry.get()
        impersonate = self.impersonate_entry.get() or None

        if not all([server, username, password]):
            self.log("Server, username, and password are required.")
            return

        self.connect_button.configure(state="disabled", text="Connecting...")
        self.log("Connecting...")

        self.worker_thread = threading.Thread(
            target=self.connect_worker,
            args=(server, username, password, impersonate)
        )
        self.worker_thread.start()

    def connect_worker(self, server, username, password, impersonate):
        result = self.exchange_service.connect(server, username, password, impersonate)
        self.log(result)
        if "successful" in result:
            self.after(0, self.post_connect_success)
        else:
            self.after(0, self.post_connect_failure)

    def post_connect_success(self):
        self.log("Connection successful. Fetching folders...")
        self.login_frame.grid_forget() # Hide login frame
        self.controls_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew") # Show controls
        self.status_label.configure(text="Status: Connected. Fetching folders...")
        self.fetch_folders()

    def post_connect_failure(self):
        self.connect_button.configure(state="normal", text="Connect")
        self.log("Connection failed.")

    def fetch_folders(self):
        self.worker_thread = threading.Thread(target=self.fetch_folders_worker)
        self.worker_thread.start()

    def fetch_folders_worker(self):
        folders = self.exchange_service.get_folders()
        self.after(0, self.populate_folder_list, folders)

    def populate_folder_list(self, folders):
        for widget in self.folder_frame.winfo_children():
            widget.destroy()

        self.selected_folder_var = ctk.StringVar()

        for i, folder in enumerate(folders):
            folder_text = f"{folder['name']} ({folder['total_count']})"
            radio = ctk.CTkRadioButton(
                self.folder_frame,
                text=folder_text,
                variable=self.selected_folder_var,
                value=str(i),
                command=lambda f=folder['obj']: self.select_folder(f)
            )
            radio.pack(anchor="w", padx=5, pady=2)

        self.status_label.configure(text="Status: Ready. Select a folder.")
        self.log("Folders loaded. Please select a folder to clean.")

    def select_folder(self, folder_obj):
        self.selected_folder_obj = folder_obj
        self.start_button.configure(state="normal")
        self.log(f"Selected folder: {folder_obj.name}")

    def start_cleaning(self):
        if not self.selected_folder_obj:
            self.log("No folder selected.")
            return

        self.start_button.configure(state="disabled")
        self.progress_bar.set(0)

        self.worker_thread = threading.Thread(
            target=self.exchange_service.empty_folder,
            args=(self.selected_folder_obj, self.update_progress)
        )
        self.worker_thread.start()

    def update_progress(self, processed, remaining, speed, est_time, status):
        # This method is called from a worker thread, so we need to schedule the GUI update
        self.after(0, self.update_progress_gui, processed, remaining, speed, est_time, status)

    def update_progress_gui(self, processed, remaining, speed, est_time, status):
        self.log(status)

        total = processed + remaining
        if total > 0:
            progress_value = processed / total
            self.progress_bar.set(progress_value)

        est_time_str = "N/A"
        if est_time > 0:
            est_time_str = time.strftime('%H:%M:%S', time.gmtime(est_time))

        status_text = f"Status: Cleaning... | Speed: {speed:.2f} items/s | Est. Time: {est_time_str}"
        self.status_label.configure(text=status_text)

        if "complete" in status or "already empty" in status:
            self.start_button.configure(state="normal")
            self.status_label.configure(text=f"Status: {status}")
            self.log("Refreshing folder list...")
            self.fetch_folders()

if __name__ == "__main__":
    app = App()
    app.mainloop()
