#!/usr/bin/env python3

import os
import getpass
import concurrent.futures
from exchangelib import Credentials, Account, Configuration, DELEGATE, EWSDateTime, EWSTimeZone
from exchangelib.errors import ErrorItemNotFound, ErrorServerBusy, ErrorMailboxStoreUnavailable
# Importer les dictionnaires de conversion de fuseaux horaires
from exchangelib.winzone import MS_TIMEZONE_TO_IANA_MAP, CLDR_TO_MS_TIMEZONE_MAP
import time
import sys
import random
import tkinter as tk
from tkinter import ttk
from datetime import datetime
import threading
import statistics
from colorama import init, Fore, Back, Style
import multiprocessing
import queue
import subprocess
import shutil
import select
import platform
import functools
import inspect

# Ajouter une entrée personnalisée pour le fuseau horaire "Customized Time Zone"
# Utiliser Europe/Paris comme valeur raisonnable (à ajuster selon vos besoins)
MS_TIMEZONE_TO_IANA_MAP['Customized Time Zone'] = "Europe/Paris"
print(f"{Fore.CYAN}Fuseau horaire personnalisé ajouté: 'Customized Time Zone' -> 'Europe/Paris'{Style.RESET_ALL}")

# Détection de la plateforme
IS_WINDOWS = False  # Force Linux mode

# Installation de rich si nécessaire
try:
    from rich.console import Console
    from rich.table import Table
    from rich.panel import Panel
    from rich.layout import Layout
    from rich.live import Live
    RICH_AVAILABLE = True
except ImportError:
    RICH_AVAILABLE = False
    print("La bibliothèque 'rich' n'est pas installée. Tentative d'installation...")
    try:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "rich"])
        print("Installation réussie! Importation de rich...")
        from rich.console import Console
        from rich.table import Table
        from rich.panel import Panel
        from rich.layout import Layout
        from rich.live import Live
        RICH_AVAILABLE = True
    except:
        print("Échec de l'installation de 'rich'. Les statistiques seront affichées en mode simple.")

# Initialize colorama
init()

# État du clignotement
blink_state = 0

# Fonction pour imprimer le logo avec alternance
def print_logo():
    """Affiche le logo du programme avec un art ASCII qui alterne"""
    global blink_state
    
    # Logo 1: EWS FOLDER
    ews_folder_logo = f"""{Fore.CYAN}╔══════════════════════════════════════════════════════════════════════════╗
{Fore.CYAN}║ {Fore.GREEN}███████╗██╗    ██╗███████╗    ███████╗ ██████╗ ██╗     ██████╗ ███████╗██████╗ {Fore.CYAN}║
{Fore.CYAN}║ {Fore.GREEN}██╔════╝██║    ██║██╔════╝    ██╔════╝██╔═══██╗██║     ██╔══██╗██╔════╝██╔══██╗{Fore.CYAN}║
{Fore.CYAN}║ {Fore.GREEN}█████╗  ██║ █╗ ██║███████╗    █████╗  ██║   ██║██║     ██║  ██║█████╗  ██████╔╝{Fore.CYAN}║
{Fore.CYAN}║ {Fore.GREEN}██╔══╝  ██║███╗██║╚════██║    ██╔══╝  ██║   ██║██║     ██║  ██║██╔══╝  ██╔══██╗{Fore.CYAN}║
{Fore.CYAN}║ {Fore.GREEN}███████╗╚███╔███╔╝███████║    ██║     ╚██████╔╝███████╗██████╔╝███████╗██║  ██║{Fore.CYAN}║
{Fore.CYAN}║ {Fore.GREEN}╚══════╝ ╚══╝╚══╝ ╚══════╝    ╚═╝      ╚═════╝ ╚══════╝╚═════╝ ╚══════╝╚═╝  ╚═╝{Fore.CYAN}║
{Fore.CYAN}║                                                                                    ║
{Fore.CYAN}║                         {Fore.YELLOW}[ EWS FOLDER CLEANER ]{Fore.CYAN}                     ║
{Fore.CYAN}╚══════════════════════════════════════════════════════════════════════════╝{Style.RESET_ALL}"""
    
    # Logo 2: EWS CLEAN
    ews_clean_logo = f"""{Fore.CYAN}╔══════════════════════════════════════════════════════════════════════════╗
{Fore.CYAN}║ {Fore.GREEN}███████╗██╗    ██╗███████╗     ██████╗██╗     ███████╗ █████╗ ███╗   ██╗{Fore.CYAN} ║
{Fore.CYAN}║ {Fore.GREEN}██╔════╝██║    ██║██╔════╝    ██╔════╝██║     ██╔════╝██╔══██╗████╗  ██║{Fore.CYAN} ║
{Fore.CYAN}║ {Fore.GREEN}█████╗  ██║ █╗ ██║███████╗    ██║     ██║     █████╗  ███████║██╔██╗ ██║{Fore.CYAN} ║
{Fore.CYAN}║ {Fore.GREEN}██╔══╝  ██║███╗██║╚════██║    ██║     ██║     ██╔══╝  ██╔══██║██║╚██╗██║{Fore.CYAN} ║
{Fore.CYAN}║ {Fore.GREEN}███████╗╚███╔███╔╝███████║    ╚██████╗███████╗███████╗██║  ██║██║ ╚████║{Fore.CYAN} ║
{Fore.CYAN}║ {Fore.GREEN}╚══════╝ ╚══╝╚══╝ ╚══════╝     ╚═════╝╚══════╝╚══════╝╚═╝  ╚═╝╚═╝  ╚═══╝{Fore.CYAN} ║
{Fore.CYAN}║                                                                                    ║
{Fore.CYAN}║                              {Fore.YELLOW}[ EWS CLEAN ]{Fore.CYAN}                          ║
{Fore.CYAN}╚══════════════════════════════════════════════════════════════════════════╝{Style.RESET_ALL}"""
    
    # Alterner entre les deux logos
    logo_to_print = ews_folder_logo if blink_state % 2 == 0 else ews_clean_logo
    blink_state += 1
    
    # Afficher le logo sélectionné
    print(logo_to_print)

# Fonction pour obtenir les identifiants
def get_credentials():
    """Get username and password interactively"""
    username = input(f"{Fore.GREEN}Enter your email address: {Style.RESET_ALL}")
    password = getpass.getpass(f"{Fore.GREEN}Enter your password: {Style.RESET_ALL}")
    return username, password

# Classe pour suivre les temps d'appel EWS
class EWSStats:
    def __init__(self):
        self.call_times = []
        self.call_types = {}  # Dictionnaire pour stocker les temps par type d'appel
        self.lock = threading.Lock()
        self.last_call_time = 0
        self.max_calls_to_keep = 1000  # Limiter le nombre d'entrées pour éviter une utilisation excessive de mémoire
        self.last_commands = {}  # Stocker les dernières commandes par type
        self.active_calls = 0  # Compte des appels actifs
    
    def add_call_time(self, ms, call_type="generic", command_details=""):
        with self.lock:
            self.last_call_time = ms
            self.call_times.append(ms)
            
            # Limiter la taille des listes
            if len(self.call_times) > self.max_calls_to_keep:
                self.call_times = self.call_times[-self.max_calls_to_keep:]
            
            # Ajouter par type
            if call_type not in self.call_types:
                self.call_types[call_type] = []
            
            self.call_types[call_type].append(ms)
            
            # Stocker les détails de la commande
            self.last_commands[call_type] = command_details
            
            # Limiter la taille des listes par type
            if len(self.call_types[call_type]) > self.max_calls_to_keep:
                self.call_types[call_type] = self.call_types[call_type][-self.max_calls_to_keep:]
            
            # Ajouter un log pour cet appel
            log_level = "INFO"
            
            # S'il s'agit d'une erreur, toujours la journaliser comme une erreur
            if call_type.startswith("error_"):
                log_level = "ERROR"
                # Forcer l'ajout de log à l'interface unifiée pour les erreurs
                if ews_unified_interface and ews_unified_interface.running:
                    ews_unified_interface.add_log(f"{command_details} - {ms:.2f}ms", log_level)
            
            # Ajouter un log spécial pour les appels lents (plus de 1000ms)
            if ms > 1000:
                # Ajouter un emoji pour rendre plus visible
                slow_log_message = f"🕒 SLOW EWS CALL: {call_type} - {ms:.2f}ms - {command_details}"
                ews_logger.add_log(slow_log_message, "WARN")
                # Également ajouter à l'interface unifiée avec priorité
                if ews_unified_interface and ews_unified_interface.running:
                    ews_unified_interface.add_log(slow_log_message, "WARN")
            
            # Log standard pour tous les appels
            log_message = f"EWS call: {command_details} - {call_type} - {ms:.2f}ms"
            ews_logger.add_log(log_message, log_level)
    
    def start_call(self):
        with self.lock:
            self.active_calls += 1
    
    def end_call(self):
        with self.lock:
            if self.active_calls > 0:
                self.active_calls -= 1
    
    def get_active_calls(self):
        with self.lock:
            return self.active_calls
    
    def get_last_command(self, call_type):
        with self.lock:
            return self.last_commands.get(call_type, "")
    
    def get_stats(self):
        with self.lock:
            if not self.call_times:
                return {"last": 0, "min": 0, "max": 0, "avg": 0, "median": 0, "count": 0, "active": self.active_calls}
            
            return {
                "last": self.last_call_time,
                "min": min(self.call_times),
                "max": max(self.call_times),
                "avg": sum(self.call_times) / len(self.call_times),
                "median": statistics.median(self.call_times) if len(self.call_times) > 0 else 0,
                "count": len(self.call_times),
                "active": self.active_calls
            }
    
    def get_type_stats(self, call_type):
        with self.lock:
            if call_type not in self.call_types or not self.call_types[call_type]:
                return {"min": 0, "max": 0, "avg": 0, "count": 0}
            
            call_list = self.call_types[call_type]
            return {
                "min": min(call_list),
                "max": max(call_list),
                "avg": sum(call_list) / len(call_list),
                "count": len(call_list),
                "last_command": self.last_commands.get(call_type, "")
            }

# Créer une instance globale pour les statistiques
ews_stats = EWSStats()

# Classe pour les logs EWS
class EWSLogger:
    def __init__(self, log_file=None):
        self.log_entries = []
        self.log_queue = queue.Queue()
        self.log_thread = None
        self.running = False
        self.log_to_console = False
        
        # Set the log file path
        if log_file:
            self.log_file = log_file
        else:
            # Default log file in /tmp on Linux or temp on Windows
            if platform.system() == "Windows":
                self.log_file = os.path.join(os.environ.get("TEMP", "C:\\Temp"), "ews_logs.txt")
            else:
                self.log_file = "/tmp/ews_logs.txt"
        
        # Ensure the log file is accessible
        self.setup_log_file()
    
    def setup_log_file(self):
        """Ensure the log file is created and accessible"""
        try:
            # Get the directory of the log file
            log_dir = os.path.dirname(self.log_file)
            
            # Check if the directory exists and create it if needed
            if not os.path.exists(log_dir) and log_dir:
                try:
                    os.makedirs(log_dir, exist_ok=True)
                    print(f"{Fore.GREEN}Created log directory: {log_dir}{Style.RESET_ALL}")
                except Exception as e:
                    print(f"{Fore.RED}Failed to create log directory: {log_dir} - {e}{Style.RESET_ALL}")
                    # Fall back to using home directory
                    home_dir = os.path.expanduser("~")
                    self.log_file = os.path.join(home_dir, "ews_logs.txt")
                    print(f"{Fore.YELLOW}Using alternate log location: {self.log_file}{Style.RESET_ALL}")
            
            # Create the log file if it doesn't exist
            with open(self.log_file, 'a') as f:
                if os.path.getsize(self.log_file) == 0:
                    # File was just created, write header
                    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    f.write(f"=== EWS CLEANER LOG STARTED AT {timestamp} ===\n")
            
            # Verify the file exists and is writeable
            if not os.path.exists(self.log_file):
                raise IOError(f"Log file could not be created: {self.log_file}")
            
            print(f"{Fore.GREEN}Log file initialized: {self.log_file}{Style.RESET_ALL}")
            print(f"{Fore.CYAN}To view logs in real-time, use: tail -f {self.log_file}{Style.RESET_ALL}")
        except Exception as e:
            print(f"{Fore.RED}Failed to initialize log file: {e}{Style.RESET_ALL}")
            # Continue without logging to file
            self.log_file = None
    
    def start_logging(self):
        """Start the logging thread"""
        pass
    
    def add_log(self, message, level="INFO"):
        """Ajoute un message au log"""
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        log_entry = {
            "timestamp": timestamp,
            "message": message,
            "level": level
        }
        # Log to console if enabled
        if self.log_to_console:
            color = Fore.GREEN
            if level == "ERROR":
                color = Fore.RED
            elif level == "WARN":
                color = Fore.YELLOW
            print(f"{color}[{timestamp}] {message}{Style.RESET_ALL}")
        
        # Store log in memory
        self.log_entries.append(log_entry)
        if len(self.log_entries) > 1000:  # limit to 1000 entries
            self.log_entries = self.log_entries[-1000:]
            
        # Log to file
        if self.log_file:
            try:
                with open(self.log_file, 'a') as f:
                    f.write(f"[{timestamp}] [{level}] {message}\n")
            except Exception as e:
                print(f"{Fore.RED}Failed to write to log file: {e}{Style.RESET_ALL}")
    
    def show_log_window(self):
        """Display log window (dummy implementation for Linux)"""
        print(f"{Fore.GREEN}Log window: using file {self.log_file}{Style.RESET_ALL}")
    
    def stop(self):
        """Stop logging"""
        self.running = False

# Classe pour afficher les statistiques EWS dans une fenêtre séparée
class EWSStatsWindow:
    def __init__(self):
        self.running = False
        self.user_exit_requested = False
    
    def show_stats_window(self):
        """Display statistics window"""
        self.running = True
        print(f"{Fore.GREEN}Statistics window opened{Style.RESET_ALL}")
    
    def exit_stats_window(self):
        """Close statistics window"""
        self.running = False
        self.user_exit_requested = True
        print(f"{Fore.YELLOW}Statistics window closed{Style.RESET_ALL}")
    
    def stop(self):
        """Stop statistics window"""
        self.running = False
        print(f"{Fore.YELLOW}Statistics window stopped{Style.RESET_ALL}")

# Modifier la classe EWSUnifiedInterface pour contrôler le moment où le monitoring démarre
class EWSUnifiedInterface:
    def __init__(self):
        self.log_queue = multiprocessing.Queue()
        self.command_queue = multiprocessing.Queue()
        self.data_queue = multiprocessing.Queue()
        self.interface_process = None
        self.running = False
        self.update_thread = None
        self.user_exit_requested = False
        self.monitoring_active = False  # Nouveau flag pour contrôler l'activation du monitoring
        # Données de progression du traitement
        self.progress_data = {
            "folder_name": "",
            "processed": 0,
            "remaining": 0,
            "speed": 0,
            "est_time": 0,
            "active": False
        }
    
    def add_log(self, message, level="INFO"):
        """Ajoute un message au log"""
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        log_entry = {
            "timestamp": timestamp,
            "message": message,
            "level": level
        }
        if self.running:
            try:
                self.log_queue.put(log_entry)
            except:
                pass
    
    def add_folders(self, folders_data):
        """Ajoute les dossiers disponibles"""
        if self.running:
            try:
                self.data_queue.put({"type": "folders", "data": folders_data})
            except:
                pass
    
    def get_command(self, timeout=0.1):
        """Récupère une commande de l'interface utilisateur"""
        try:
            return self.command_queue.get(timeout=timeout)
        except queue.Empty:
            return None
    
    def start(self, start_monitoring=False):
        """Démarre l'interface unifiée"""
        self.running = True
        self.monitoring_active = start_monitoring
        print(f"{Fore.GREEN}Interface unifiée démarrée (monitoring: {start_monitoring}){Style.RESET_ALL}")
    
    def start_monitoring(self):
        """Active le monitoring des appels EWS"""
        if not self.monitoring_active and self.running:
            self.monitoring_active = True
            stop_event = multiprocessing.Event()
            
            self.interface_process = multiprocessing.Process(
                target=interface_process,
                args=(self.command_queue, self.data_queue, self.log_queue, stop_event)
            )
            self.interface_process.daemon = True
            self.interface_process.start()
            
            # Démarrer le thread de mise à jour des statistiques
            self.update_thread = threading.Thread(target=self.update_stats)
            self.update_thread.daemon = True
            self.update_thread.start()
            
            print(f"{Fore.GREEN}Monitoring EWS activé{Style.RESET_ALL}")
            return True
        return False
    
    def update_stats(self):
        """Met à jour périodiquement les statistiques"""
        last_stats = {}
        
        while self.running and self.monitoring_active:
            try:
                # Récupérer les statistiques actuelles
                stats = ews_stats.get_stats()
                
                # Ne mettre à jour que si les stats ont changé
                if stats != last_stats:
                    last_stats = stats.copy()
                    
                    # Préparer les données de type d'appel
                    call_types = {}
                    for call_type in ews_stats.call_types.keys():
                        call_types[call_type] = ews_stats.get_type_stats(call_type)
                    
                    # Envoyer les stats au processus d'interface
                    stats_data = {
                        "type": "stats",
                        "data": {
                            "stats": stats,
                            "call_types": call_types,
                            "timestamp": datetime.now().strftime("%H:%M:%S")
                        }
                    }
                    
                    if self.running and self.monitoring_active:
                        self.data_queue.put(stats_data)
            
            except Exception as e:
                print(f"Error updating stats: {e}")
            
            # Attendre avant la prochaine mise à jour
            time.sleep(0.5)
    
    def stop(self):
        """Arrête proprement l'interface"""
        self.running = False
        self.monitoring_active = False
        
        if self.update_thread and self.update_thread.is_alive():
            self.update_thread.join(timeout=2.0)
        
        if self.interface_process and self.interface_process.is_alive():
            try:
                self.interface_process.terminate()
            except:
                pass
        
        print(f"{Fore.YELLOW}Interface unifiée arrêtée{Style.RESET_ALL}")

    def update_progress(self, folder_name, processed, remaining, speed, est_time):
        """Met à jour les données de progression"""
        self.progress_data = {
            "folder_name": folder_name,
            "processed": processed,
            "remaining": remaining,
            "speed": speed,
            "est_time": est_time,
            "active": True
        }
        
        # Envoyer les données de progression au processus d'interface
        if self.running and self.monitoring_active:
            try:
                self.data_queue.put({
                    "type": "progress",
                    "data": self.progress_data
                })
            except:
                pass
    
    def reset_progress(self):
        """Réinitialise les données de progression"""
        self.progress_data = {
            "folder_name": "",
            "processed": 0,
            "remaining": 0,
            "speed": 0,
            "est_time": 0,
            "active": False
        }
        
        # Envoyer les données de progression réinitialisées
        if self.running and self.monitoring_active:
            try:
                self.data_queue.put({
                    "type": "progress",
                    "data": self.progress_data
                })
            except:
                pass

# Fonction pour intercepter et chronométrer les appels EWS
def intercept_ews_calls():
    """Intercepte les appels EWS pour mesurer leur temps d'exécution"""
    from exchangelib.protocol import Protocol
    
    # Essayer d'intercepter la méthode post (qui est généralement celle qui fait les requêtes HTTP)
    if hasattr(Protocol, 'post'):
        # Sauvegarder la méthode originale
        original_post = Protocol.post
        
        @functools.wraps(original_post)
        def wrapped_post(self, *args, **kwargs):
            # Mesurer le temps d'exécution
            ews_stats.start_call()
            start_time = time.time()
            
            # Essayer de déterminer le type d'appel
            call_type = "unknown"
            call_info = "EWS request"
            
            # Dans les nouveaux arguments, on trouve généralement 'data' qui contient les info de la requête
            try:
                if len(args) > 0 and hasattr(args[0], 'tag'):
                    call_type = args[0].tag.localname
                    call_info = f"Operation: {call_type}"
                elif 'data' in kwargs and hasattr(kwargs['data'], 'tag'):
                    call_type = kwargs['data'].tag.localname
                    call_info = f"Operation: {call_type}"
            except:
                pass
            
            try:
                response = original_post(self, *args, **kwargs)
                elapsed_ms = (time.time() - start_time) * 1000
                ews_stats.add_call_time(elapsed_ms, call_type, call_info)
                ews_stats.end_call()
                return response
            except Exception as e:
                elapsed_ms = (time.time() - start_time) * 1000
                error_message = f"Erreur EWS ({call_type}): {str(e)}"
                
                # Pour les appels lents mais réussis
                if elapsed_ms > 1000:
                    pause_duration = 3  # pause plus courte pour les appels réussis
                    current_time = datetime.now().strftime("%H:%M:%S")
                    
                    # Message de pause avec horodatage
                    pause_message = f"⚠️ DÉBUT PAUSE DE {pause_duration}s ({current_time}): Opération lente mais réussie - {elapsed_ms:.2f}ms"
                    print(f"{Fore.YELLOW}{pause_message}{Style.RESET_ALL}")
                    ews_logger.add_log(pause_message, "WARN")
                    if ews_unified_interface and ews_unified_interface.running:
                        ews_unified_interface.add_log(pause_message, "WARN")
                    
                    # Exécuter la pause avec vérification
                    pause_start = time.time()
                    time.sleep(pause_duration)
                    actual_pause = time.time() - pause_start
                    
                    # Message après la pause
                    current_time = datetime.now().strftime("%H:%M:%S")
                    end_pause_message = f"✅ FIN PAUSE ({current_time}): Durée réelle = {actual_pause:.1f}s"
                    print(f"{Fore.GREEN}{end_pause_message}{Style.RESET_ALL}")
                    ews_logger.add_log(end_pause_message, "INFO")
                    if ews_unified_interface and ews_unified_interface.running:
                        ews_unified_interface.add_log(end_pause_message, "INFO")
                
                # Pour les erreurs de base de données
                if "mailbox database is temporarily unavailable" in str(e).lower():
                    pause_duration = 30  # 30 secondes de pause
                    current_time = datetime.now().strftime("%H:%M:%S")
                    
                    # Message AVANT la pause, avec horodatage
                    pause_message = f"⚠️⚠️ DÉBUT PAUSE DE {pause_duration}s ({current_time}): Base de données indisponible - {elapsed_ms:.2f}ms"
                    print(f"{Fore.RED}{pause_message}{Style.RESET_ALL}")
                    # Ajouter le message de pause aux logs et à l'interface avec haute priorité
                    ews_logger.add_log(pause_message, "ERROR")
                    ews_unified_interface.add_log(pause_message, "ERROR")
                    
                    # Exécuter la pause avec vérification
                    pause_start = time.time()
                    time.sleep(pause_duration)
                    actual_pause = time.time() - pause_start
                    
                    # Message APRÈS la pause, avec durée réelle
                    current_time = datetime.now().strftime("%H:%M:%S")
                    end_pause_message = f"✅ FIN PAUSE ({current_time}): Durée réelle = {actual_pause:.1f}s"
                    print(f"{Fore.GREEN}{end_pause_message}{Style.RESET_ALL}")
                    ews_logger.add_log(end_pause_message, "INFO")
                    ews_unified_interface.add_log(end_pause_message, "INFO")
                
                ews_stats.add_call_time(elapsed_ms, f"error_{call_type}", f"Error in {call_info}: {str(e)}")
                ews_logger.add_log(error_message, "ERROR")
                ews_stats.end_call()
                raise
        
        # Remplacer la méthode originale par notre wrapper
        Protocol.post = wrapped_post
        print(f"{Fore.GREEN}Monitoring EWS activé: méthode Protocol.post interceptée{Style.RESET_ALL}")
        return True
    
    # Si post n'existe pas, essayer avec send qui est aussi couramment utilisé
    elif hasattr(Protocol, 'send'):
        # Sauvegarder la méthode originale
        original_send = Protocol.send
        
        @functools.wraps(original_send)
        def wrapped_send(self, *args, **kwargs):
            # Mesurer le temps d'exécution
            ews_stats.start_call()
            start_time = time.time()
            
            # Essayer de déterminer le type d'appel
            call_type = "unknown"
            call_info = "EWS request"
            
            try:
                if len(args) > 0 and hasattr(args[0], 'tag'):
                    call_type = args[0].tag.localname
                    call_info = f"Operation: {call_type}"
                elif 'data' in kwargs and hasattr(kwargs['data'], 'tag'):
                    call_type = kwargs['data'].tag.localname
                    call_info = f"Operation: {call_type}"
            except:
                pass
            
            try:
                response = original_send(self, *args, **kwargs)
                elapsed_ms = (time.time() - start_time) * 1000
                ews_stats.add_call_time(elapsed_ms, call_type, call_info)
                ews_stats.end_call()
                return response
            except Exception as e:
                elapsed_ms = (time.time() - start_time) * 1000
                error_message = f"Erreur EWS ({call_type}): {str(e)}"
                ews_stats.add_call_time(elapsed_ms, f"error_{call_type}", f"Error in {call_info}: {str(e)}")
                ews_logger.add_log(error_message, "ERROR")
                ews_stats.end_call()
                raise
        
        # Remplacer la méthode originale par notre wrapper
        Protocol.send = wrapped_send
        print(f"{Fore.GREEN}Monitoring EWS activé: méthode Protocol.send interceptée{Style.RESET_ALL}")
        return True
    
    # Si aucune méthode standard n'est trouvée, essayer une autre approche - intercepter les appels au niveau de Account
    from exchangelib.items import Item
    
    if hasattr(Item, 'delete'):
        original_delete = Item.delete
        
        @functools.wraps(original_delete)
        def wrapped_delete(self, *args, **kwargs):
            # Mesurer le temps d'exécution
            ews_stats.start_call()
            start_time = time.time()
            call_type = "delete"
            call_info = f"Delete item: {self.__class__.__name__}"
            
            try:
                response = original_delete(self, *args, **kwargs)
                elapsed_ms = (time.time() - start_time) * 1000
                ews_stats.add_call_time(elapsed_ms, call_type, call_info)
                ews_stats.end_call()
                
                # Ajouter une pause pour les appels réussis mais lents
                if elapsed_ms > 1000:
                    pause_duration = 3  # pause plus courte pour les appels réussis
                    time_before = datetime.now().strftime("%H:%M:%S.%f")[:-3]
                    pause_message = f"⏱️ DÉBUT PAUSE {time_before} - Suppression lente mais réussie ({elapsed_ms:.0f}ms) - {pause_duration}s d'attente"
                    print(f"{Fore.YELLOW}{pause_message}{Style.RESET_ALL}")
                    ews_logger.add_log(pause_message, "WARN")
                    ews_unified_interface.add_log(pause_message, "WARN")
                    
                    time.sleep(pause_duration)
                    
                    time_after = datetime.now().strftime("%H:%M:%S.%f")[:-3]
                    after_message = f"✅ FIN PAUSE {time_after} - Reprise après {pause_duration}s d'attente"
                    print(f"{Fore.GREEN}{after_message}{Style.RESET_ALL}")
                    ews_logger.add_log(after_message, "INFO")
                    ews_unified_interface.add_log(after_message, "INFO")
                
                return response
            except Exception as e:
                elapsed_ms = (time.time() - start_time) * 1000
                error_message = f"Erreur lors de la suppression d'un élément: {str(e)}"
                error_type = "error_delete"
                
                # Log amélioré pour les erreurs avec haute latence
                if elapsed_ms > 1000:
                    error_type = "error_delete_slow"
                    ews_logger.add_log(f"ERREUR LENTE DELETE: {error_message} - {elapsed_ms:.2f}ms", "ERROR")
                    ews_unified_interface.add_log(f"Erreur de suppression lente détectée: {elapsed_ms:.2f}ms", "ERROR")
                    
                    # Pause forcée pour les suppressions lentes
                    if "mailbox database is temporarily unavailable" in str(e).lower():
                        pause_duration = 30  # 30 secondes de pause
                        time_before = datetime.now().strftime("%H:%M:%S.%f")[:-3]
                        pause_message = f"⛔ DÉBUT PAUSE {time_before} - Base de données indisponible ({elapsed_ms:.0f}ms) - {pause_duration}s d'attente"
                        print(f"{Fore.RED}{pause_message}{Style.RESET_ALL}")
                        ews_logger.add_log(pause_message, "ERROR")
                        ews_unified_interface.add_log(pause_message, "ERROR")
                        
                        time.sleep(pause_duration)
                        
                        time_after = datetime.now().strftime("%H:%M:%S.%f")[:-3]
                        after_message = f"✅ FIN PAUSE {time_after} - Reprise après {pause_duration}s d'attente"
                        print(f"{Fore.GREEN}{after_message}{Style.RESET_ALL}")
                        ews_logger.add_log(after_message, "INFO")
                        ews_unified_interface.add_log(after_message, "INFO")
                    else:
                        # Autres erreurs lentes
                        pause_duration = 5  # 5 secondes de pause
                        pause_message = f"⚠️ PAUSE DE {pause_duration}s: Autre erreur de suppression lente - {elapsed_ms:.2f}ms"
                        print(f"{Fore.YELLOW}{pause_message}{Style.RESET_ALL}")
                        ews_logger.add_log(pause_message, "WARN")
                        ews_unified_interface.add_log(pause_message, "WARN")
                        time.sleep(pause_duration)
                
                ews_stats.add_call_time(elapsed_ms, error_type, f"Error in {call_info}: {str(e)}")
                ews_logger.add_log(error_message, "ERROR")
                ews_stats.end_call()
                raise
        
        # Remplacer la méthode originale par notre wrapper
        Item.delete = wrapped_delete
        print(f"{Fore.GREEN}Monitoring EWS activé: méthode Item.delete interceptée{Style.RESET_ALL}")
        return True
    
    # Si aucune méthode n'a pu être interceptée
    print(f"{Fore.RED}Impossible d'activer le monitoring EWS: aucune méthode connue n'a été trouvée dans exchangelib{Style.RESET_ALL}")
    return False

# Fonction pour lister les dossiers
def list_folders(account):
    """List all folders in the account"""
    print(f"\n{Fore.CYAN}Listing folders...{Style.RESET_ALL}")
    folders = []
    
    try:
        # Parcourir tous les dossiers récursivement
        for i, folder in enumerate(account.root.walk(), 1):
            try:
                # Ajouter des informations sur le dossier
                print(f"{Fore.WHITE}{i}. {Fore.GREEN}{folder.name} {Fore.CYAN}({folder.total_count} items){Style.RESET_ALL}")
                folders.append(folder)
                
                # Enregistrer dans les logs
                ews_logger.add_log(f"Found folder: {folder.name} ({folder.total_count} items)")
                
                # Nous n'avons pas besoin d'ajouter manuellement des mesures de temps ici
                # car l'intercepteur les mesure déjà
            except Exception as e:
                print(f"{Fore.RED}Error getting folder info: {e}{Style.RESET_ALL}")
                ews_logger.add_log(f"Error getting folder info: {e}", "ERROR")
    except Exception as e:
        print(f"{Fore.RED}Error listing folders: {e}{Style.RESET_ALL}")
        ews_logger.add_log(f"Error listing folders: {e}", "ERROR")
    
    return folders

# Fonction pour vider un dossier
def empty_folder(folder, batch_size=100):
    """Empty a folder by deleting all items in batches"""
    print(f"\n{Fore.YELLOW}Processing folder: {Fore.GREEN}{folder.name}{Style.RESET_ALL}")
    print(f"{Fore.CYAN}Total items: {folder.total_count}{Style.RESET_ALL}")
    
    if folder.total_count == 0:
        print(f"{Fore.GREEN}Folder is already empty.{Style.RESET_ALL}")
        return
    
    # Commencer à mesurer le temps
    start_time = time.time()
    processed = 0
    
    try:
        # Récupérer l'ID du dossier pour pouvoir le retrouver après
        folder_id = folder.id
        folder_name = folder.name
        folder_path = folder.absolute
        account = folder.account
        
        # Traiter par lots
        while True:
            # Récupérer le dossier actualisé depuis le serveur
            try:
                # ... code existant pour rafraîchir le dossier ...
                if hasattr(folder, 'refresh'):
                    folder.refresh()
                    actual_remaining = folder.total_count
                else:
                    # ... code existant ...
                    if hasattr(account.root, 'get_folder_by_path'):
                        try:
                            folder = account.root.get_folder_by_path(folder_path)
                            actual_remaining = folder.total_count
                        except:
                            actual_remaining = folder.total_count
                    else:
                        actual_remaining = folder.total_count
                
                if actual_remaining == 0:
                    print(f"{Fore.GREEN}Folder is now empty.{Style.RESET_ALL}")
                    # Réinitialiser les données de progression
                    ews_unified_interface.reset_progress()
                    break
            except Exception as e:
                print(f"{Fore.RED}Error refreshing folder info: {e}{Style.RESET_ALL}")
                ews_logger.add_log(f"Error refreshing folder info: {e}", "ERROR")
                actual_remaining = folder.total_count
                if actual_remaining == 0:
                    ews_unified_interface.reset_progress()
                    break
            
            # Récupérer un lot d'éléments
            # Nous n'avons pas besoin de mesurer manuellement car l'intercepteur le fait
            try:
                items = list(folder.all().order_by('datetime_received')[:batch_size])
            except Exception as e:
                error_message = str(e)
                print(f"{Fore.RED}Error getting items: {error_message}{Style.RESET_ALL}")
                
                # Vérifier si c'est l'erreur spécifique de base de données temporairement indisponible
                if "mailbox database is temporarily unavailable" in error_message.lower():
                    wait_time = 30  # 30 secondes d'attente pour cette erreur spécifique
                    error_log = f"Base de données boîte aux lettres temporairement indisponible. Attente de {wait_time} secondes."
                    print(f"{Fore.YELLOW}{error_log}{Style.RESET_ALL}")
                    ews_logger.add_log(error_log, "WARN")
                    time.sleep(wait_time)
                else:
                    # Pour les autres erreurs, attendre seulement 2 secondes
                    ews_logger.add_log(f"Error getting items: {error_message}", "ERROR")
                    time.sleep(2)  # Pause en cas d'erreur
                
                continue
            
            if not items:
                print(f"{Fore.YELLOW}No more items found in folder.{Style.RESET_ALL}")
                ews_unified_interface.reset_progress()
                break
            
            # Supprimer le lot
            print(f"{Fore.CYAN}Deleting batch of {len(items)} items...{Style.RESET_ALL}")
            ews_logger.add_log(f"Deleting batch of {len(items)} items from {folder.name}")
            
            # Ajouter un suivi des échecs consécutifs
            consecutive_failures = 0
            
            try:
                # Supprimer les éléments un par un - l'intercepteur mesurera chaque appel
                for item in items:
                    try:
                        # Ajouter un délai entre chaque appel pour réduire la charge
                        time.sleep(0.05)  # 50ms entre chaque appel
                        
                        item.delete()  # L'intercepteur mesure automatiquement cet appel
                        consecutive_failures = 0  # Réinitialiser le compteur en cas de succès
                    except Exception as delete_error:
                        error_message = str(delete_error)
                        
                        # Vérifier si c'est l'erreur spécifique de base de données temporairement indisponible
                        # mais seulement si elle n'a pas déjà été traitée par le wrapped_delete
                        if "mailbox database is temporarily unavailable" in error_message.lower() and "⚠️ PAUSE" not in error_message:
                            consecutive_failures += 1
                            
                            # Si nous avons plusieurs échecs consécutifs, prendre une pause plus longue
                            if consecutive_failures >= 3:
                                pause_duration = 60  # 1 minute de pause
                                pause_message = f"⚠️ PAUSE DE {pause_duration}s: Base de données temporairement indisponible (échecs multiples)"
                                print(f"{Fore.RED}{pause_message}{Style.RESET_ALL}")
                                ews_logger.add_log(pause_message, "ERROR")
                                
                                # Ajouter un log dans l'interface unifiée
                                ews_unified_interface.add_log(pause_message, "ERROR")
                                
                                # Pause
                                time.sleep(pause_duration)
                                consecutive_failures = 0  # Réinitialiser après la pause
                            else:
                                # Pause plus courte pour les premières erreurs
                                short_pause = 5  # 5 secondes
                                print(f"{Fore.YELLOW}Base de données temporairement indisponible. Pause de {short_pause} secondes.{Style.RESET_ALL}")
                                ews_logger.add_log(f"Base de données temporairement indisponible. Pause de {short_pause} secondes.", "WARN")
                                time.sleep(short_pause)
                        else:
                            error_message = f"Erreur de suppression EWS: {str(delete_error)}"
                            print(f"{Fore.RED}{error_message}{Style.RESET_ALL}")
                            ews_logger.add_log(error_message, "ERROR")
                            # Continuer avec les autres éléments
                
                processed += len(items)
                
                # Mettre à jour le compteur sans essayer de rafraîchir le dossier
                try:
                    actual_remaining = folder.total_count 
                except Exception as e:
                    print(f"{Fore.YELLOW}Could not get folder count: {e}{Style.RESET_ALL}")
                    actual_remaining = max(0, folder.total_count - len(items))
                
                # Afficher la progression
                elapsed = time.time() - start_time
                items_per_sec = processed / elapsed if elapsed > 0 else 0
                est_time = actual_remaining / items_per_sec if items_per_sec > 0 else 0
                
                # Mettre à jour les données de progression dans l'interface
                ews_unified_interface.update_progress(
                    folder_name=folder.name,
                    processed=processed,
                    remaining=actual_remaining,
                    speed=items_per_sec,
                    est_time=est_time
                )
                
                formatted_time = format_time_remaining(est_time)
                print(f"{Fore.GREEN}Processed: {processed} items, {Fore.YELLOW}Remaining: {actual_remaining}, {Fore.CYAN}Speed: {items_per_sec:.2f} items/sec, {Fore.MAGENTA}Est. time left: {formatted_time}{Style.RESET_ALL}")
                
                # Ajout d'un délai pour éviter trop de charge sur le serveur
                time.sleep(0.1)
            
            except ErrorServerBusy as e:
                back_off = e.back_off
                print(f"{Fore.YELLOW}Server busy. Waiting for {back_off} seconds...{Style.RESET_ALL}")
                ews_logger.add_log(f"Server busy. Waiting for {back_off} seconds", "WARN")
                time.sleep(back_off)
            
            except ErrorMailboxStoreUnavailable:
                print(f"{Fore.YELLOW}Mailbox store unavailable. Waiting...{Style.RESET_ALL}")
                ews_logger.add_log("Mailbox store unavailable. Waiting...", "WARN")
                time.sleep(5)
            
            except Exception as e:
                print(f"{Fore.RED}Error deleting items: {e}{Style.RESET_ALL}")
                ews_logger.add_log(f"Error deleting items: {e}", "ERROR")
                time.sleep(2)  # Pause en cas d'erreur
        
        # Afficher le résumé
        elapsed = time.time() - start_time
        print(f"\n{Fore.GREEN}Folder processing complete!{Style.RESET_ALL}")
        print(f"{Fore.CYAN}Processed {processed} items in {elapsed:.2f} seconds ({processed/elapsed:.2f} items/sec){Style.RESET_ALL}")
        ews_logger.add_log(f"Completed processing folder {folder.name}. Deleted {processed} items in {elapsed:.2f} seconds.")
        
        # Réinitialiser les données de progression après avoir terminé
        ews_unified_interface.reset_progress()
    
    except Exception as e:
        print(f"{Fore.RED}Error processing folder: {e}{Style.RESET_ALL}")
        ews_logger.add_log(f"Error processing folder: {e}", "ERROR")
        # Réinitialiser les données de progression en cas d'erreur
        ews_unified_interface.reset_progress()

# Créer des instances globales pour les logs et les statistiques
ews_logger = EWSLogger()
ews_stats_window = EWSStatsWindow()

# Créer une instance globale de l'interface unifiée
ews_unified_interface = EWSUnifiedInterface()

# Fonction pour formater le temps en jours, heures, minutes, secondes
def format_time_remaining(seconds):
    """Convertit un temps en secondes en format jours, heures, minutes, secondes"""
    if seconds <= 0:
        return "0s"
    
    days, remainder = divmod(int(seconds), 86400)
    hours, remainder = divmod(remainder, 3600)
    minutes, seconds = divmod(remainder, 60)
    
    result = ""
    if days > 0:
        result += f"{days}j "
    if hours > 0 or days > 0:
        result += f"{hours}h "
    if minutes > 0 or hours > 0 or days > 0:
        result += f"{minutes}m "
    result += f"{seconds}s"
    
    return result

# Fonction pour gérer l'interface dans un processus séparé
def interface_process(command_queue, data_queue, log_queue, stop_event):
    """Processus séparé qui gère l'interface utilisateur avec un affichage console simple"""
    try:
        # État de l'interface
        logs = []
        stats_data = None
        ews_calls = []  # Pour stocker les appels EWS récents
        max_calls_to_display = 20  # Nombre maximum d'appels à afficher
        progress_data = None  # Pour stocker les informations de progression actuelle
        progress_history = []  # Pour stocker l'historique des progressions
        last_progress_time = 0  # Pour limiter la fréquence d'ajout à l'historique
        
        print("\033[92m=== EWS Cleaner - Linux Edition ===\033[0m")
        print("\033[36mMonitoring démarré. Appuyez sur Ctrl+C pour arrêter.\033[0m")
        print("-" * 80)
        
        # Boucle principale
        while not stop_event.is_set():
            try:
                # Vérifier les données
                if not data_queue.empty():
                    data = data_queue.get_nowait()
                    
                    if data["type"] == "stats":
                        stats_data = data
                        
                        # Extraire les informations d'appel EWS
                        if "call_types" in data["data"]:
                            for call_type, stats in data["data"]["call_types"].items():
                                if "last_command" in stats and stats["last_command"]:
                                    timestamp = datetime.now().strftime("%H:%M:%S")
                                    ews_calls.append({
                                        "timestamp": timestamp,
                                        "type": call_type,
                                        "time": stats["avg"],
                                        "details": stats["last_command"]
                                    })
                    elif data["type"] == "progress":
                        # Mettre à jour les données de progression
                        progress_data = data["data"]
                        
                        # Ajouter à l'historique de progression si c'est une progression active
                        # et qu'au moins 5 secondes se sont écoulées depuis la dernière entrée
                        current_time = time.time()
                        if progress_data and progress_data["active"] and (current_time - last_progress_time >= 5):
                            progress_entry = progress_data.copy()
                            progress_entry["timestamp"] = datetime.now().strftime("%H:%M:%S")
                            progress_history.append(progress_entry)
                            last_progress_time = current_time
                            
                            # Limiter la taille de l'historique
                            if len(progress_history) > 20:  # Conserver les 20 dernières entrées
                                progress_history = progress_history[-20:]
                
                # Vérifier les logs
                if not log_queue.empty():
                    log = log_queue.get_nowait()
                    logs.append(log)
                    
                    # Si c'est un log d'appel EWS, l'ajouter aux appels
                    if "EWS call" in log["message"]:
                        try:
                            parts = log["message"].split(" - ")
                            if len(parts) >= 3:
                                call_type = parts[1].strip()
                                time_str = parts[2].strip().replace("ms", "").strip()
                                time_ms = float(time_str)
                                details = parts[0] if len(parts) == 3 else parts[3]
                                
                                ews_calls.append({
                                    "timestamp": log["timestamp"],
                                    "type": call_type,
                                    "time": time_ms,
                                    "details": details
                                })
                        except:
                            pass
                
                # Limiter les collections
                if len(logs) > 100:
                    logs = logs[-100:]
                if len(ews_calls) > max_calls_to_display * 2:
                    ews_calls = ews_calls[-max_calls_to_display:]
                
                # Afficher les statistiques et derniers appels toutes les 2 secondes
                if stats_data and (time.time() % 2) < 0.1:
                    # Effacer l'écran (compatible Linux)
                    print("\033c", end="")
                    
                    # Afficher le logo
                    print_logo()
                    
                    # Date et heure actuelles
                    current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    print(f"\033[97m{current_time:^80}\033[0m")
                    print("-" * 80)
                    
                    # Afficher les informations de progression si disponibles (progression actuelle)
                    if progress_data and progress_data["active"]:
                        print("\033[93m=== PROGRESSION DU TRAITEMENT ACTUEL ===\033[0m")
                        print(f"Dossier: \033[96m{progress_data['folder_name']}\033[0m")
                        formatted_time = format_time_remaining(progress_data['est_time'])
                        print(f"Traités: \033[92m{progress_data['processed']}\033[0m éléments | " +
                              f"Restants: \033[93m{progress_data['remaining']}\033[0m | " +
                              f"Vitesse: \033[96m{progress_data['speed']:.2f}\033[0m items/sec | " +
                              f"Temps restant: \033[95m{formatted_time}\033[0m")
                        print("-" * 80)
                    
                    # Afficher l'historique des progressions
                    if progress_history:
                        print("\033[93m=== HISTORIQUE DE PROGRESSION ===\033[0m")
                        print(f"{'Heure':<10} {'Dossier':<20} {'Traités':<10} {'Restants':<10} {'Vitesse/s':<10} {'Temps rest.':<10}")
                        print("-" * 80)
                        
                        # Afficher les 5 dernières entrées d'historique en ordre inverse (le plus récent en haut)
                        for entry in reversed(progress_history[-5:]):
                            formatted_time = format_time_remaining(entry['est_time'])
                            print(f"{entry['timestamp']:<10} " +
                                 f"\033[96m{entry['folder_name'][:20]:<20}\033[0m " +
                                 f"\033[92m{entry['processed']:<10}\033[0m " +
                                 f"\033[93m{entry['remaining']:<10}\033[0m " +
                                 f"\033[96m{entry['speed']:.2f}\033[0m".ljust(11) +
                                 f"\033[95m{formatted_time}\033[0m")
                        print("-" * 80)
                    
                    # Afficher les statistiques
                    stats = stats_data["data"]["stats"]
                    print("\033[96m--- Statistiques ---\033[0m")
                    print(f"Appels actifs: \033[92m{stats['active']}\033[0m | Total: \033[92m{stats['count']}\033[0m")
                    
                    if stats['count'] > 0:
                        min_time = stats['min']
                        avg_time = stats['avg']
                        max_time = stats['max']
                        
                        # Coloriser selon les temps
                        min_color = "\033[92m"  # Vert
                        avg_color = "\033[92m"  # Vert par défaut
                        max_color = "\033[92m"  # Vert par défaut
                        
                        if avg_time > 1000:
                            avg_color = "\033[91m"  # Rouge
                        elif avg_time > 500:
                            avg_color = "\033[93m"  # Jaune
                            
                        if max_time > 1000:
                            max_color = "\033[91m"  # Rouge
                        elif max_time > 500:
                            max_color = "\033[93m"  # Jaune
                        
                        print(f"Min: {min_color}{min_time:.2f}ms\033[0m | " +
                              f"Avg: {avg_color}{avg_time:.2f}ms\033[0m | " +
                              f"Max: {max_color}{max_time:.2f}ms\033[0m")
                    
                    # Afficher les derniers logs
                    print("\n\033[96m--- Logs récents ---\033[0m")
                    for log in logs[-5:]:  # Limiter à 5 derniers logs
                        color = "\033[92m"  # Vert par défaut
                        if log["level"] == "ERROR":
                            color = "\033[91m"  # Rouge
                        elif log["level"] == "WARN":
                            color = "\033[93m"  # Jaune
                        
                        print(f"[{log['timestamp']}] {color}{log['message']}\033[0m")
                    
                    # Afficher les appels EWS récents
                    print("\n\033[96m--- Monitoring des requêtes EWS ---\033[0m")
                    print(f"{'Heure':<10} {'Type':<15} {'Temps (ms)':<12} {'Détails':<200}")
                    print("-" * 240)
                    
                    # Limiter le nombre d'erreurs de chaque type pour éviter qu'elles restent affichées en permanence
                    error_counts = {}  # Pour compter les erreurs de chaque type
                    filtered_calls = []
                    
                    # Parcourir les appels récents et limiter les erreurs
                    for call in ews_calls[-30:]:  # Considérer les 30 derniers appels
                        call_type = call["type"]
                        
                        # Si c'est une erreur, ne garder que la plus récente de chaque type
                        if call_type.startswith("error_"):
                            if call_type not in error_counts:
                                error_counts[call_type] = 0
                            
                            # Ne garder qu'une seule erreur de chaque type
                            if error_counts[call_type] < 1:
                                filtered_calls.append(call)
                                error_counts[call_type] += 1
                        else:
                            # Pour les appels normaux, les ajouter tous
                            filtered_calls.append(call)
                    
                    # Ne garder que les 10 appels les plus récents après filtrage
                    recent_calls = sorted(filtered_calls[-10:], key=lambda x: x.get('timestamp', ''))
                    
                    for call in reversed(recent_calls):  # Afficher les plus récents en premier
                        # Coloriser selon le temps
                        time_color = "\033[92m"  # Vert
                        if call["time"] > 1000:
                            time_color = "\033[91m"  # Rouge
                        elif call["time"] > 500:
                            time_color = "\033[93m"  # Jaune
                        
                        # Tronquer les détails s'ils sont trop longs
                        details = call["details"]
                        if len(details) > 200:  # Augmenté à 200 caractères
                            details = details[:197] + "..."
                        
                        print(f"{call['timestamp']:<10} " +
                              f"\033[96m{call['type'][:15]:<15}\033[0m " +
                              f"{time_color}{call['time']:.2f}ms\033[0m ".ljust(12) +
                              f"{details:<200}")  # Augmenté à 200 caractères
                    
                    print("\n\033[36mAppuyez sur Ctrl+C pour arrêter le monitoring.\033[0m")
                
                # Petite pause pour éviter de consommer trop de CPU
                time.sleep(0.1)
                
            except queue.Empty:
                pass
            except KeyboardInterrupt:
                break
            except Exception as e:
                print(f"\033[91mErreur dans l'interface: {e}\033[0m")
        
        print("\033[92mMonitoring terminé.\033[0m")
    except Exception as e:
        print(f"\033[91mErreur dans le processus d'interface: {e}\033[0m")
        import traceback
        traceback.print_exc()

# Modifier la fonction principale pour mettre en place le nouveau flux
def main():
    # S'assurer que le multiprocessing est correctement initialisé
    multiprocessing.freeze_support()
    
    # Afficher le logo
    print_logo()
    
    # Options en ligne de commande améliorées
    if len(sys.argv) > 1:
        # Vérifier si --help est demandé
        if sys.argv[1] in ['-h', '--help']:
            print(f"{Fore.GREEN}Usage: {sys.argv[0]} [OPTIONS] [username] [password] [impersonated_user]{Style.RESET_ALL}")
            print(f"{Fore.CYAN}Options:{Style.RESET_ALL}")
            print(f"  -h, --help          Affiche cette aide")
            print(f"  --no-log-window     Désactive l'ouverture d'une fenêtre de log séparée")
            print(f"  --console-log       Active l'affichage des logs dans la console principale")
            print(f"  --no-stats-window   Désactive l'ouverture d'une fenêtre de statistiques séparée")
            print(f"  --server SERVER     Spécifie le serveur Exchange")
            print(f"  --classic-ui        Utilise l'interface classique au lieu de l'interface rich unifiée")
            print(f"  --auto-monitor      Démarre automatiquement le monitoring EWS (sinon: activation manuelle)")
            sys.exit(0)
    
    # Traiter les options
    server = ''  # Serveur à spécifier obligatoirement
    show_log_window = True
    show_stats_window = True
    console_log = False
    use_unified_interface = True
    auto_monitor = False  # Par défaut, ne pas démarrer le monitoring automatiquement
    
    # Filtrer les arguments pour extraire les options
    filtered_args = []
    i = 1
    while i < len(sys.argv):
        if sys.argv[i] == '--no-log-window':
            show_log_window = False
        elif sys.argv[i] == '--no-stats-window':
            show_stats_window = False
        elif sys.argv[i] == '--console-log':
            console_log = True
        elif sys.argv[i] == '--classic-ui':
            use_unified_interface = False
        elif sys.argv[i] == '--auto-monitor':
            auto_monitor = True
        elif sys.argv[i] == '--server' and i+1 < len(sys.argv):
            server = sys.argv[i+1]
            i += 1
        else:
            filtered_args.append(sys.argv[i])
        i += 1
    
    # Créer et configurer le logger EWS
    ews_logger.log_to_console = console_log
    
    # Afficher la fenêtre de logs EWS si demandé
    if show_log_window:
        ews_logger.show_log_window()
    
    # Afficher la fenêtre de statistiques si demandé
    if show_stats_window:
        ews_stats_window.show_stats_window()
    
    # Vérifier les arguments de connexion
    if len(filtered_args) >= 2:
        username = filtered_args[0]
        password = filtered_args[1]
        impersonated_user = filtered_args[2] if len(filtered_args) > 2 else None
    else:
        # Get credentials interactively if not provided via command line
        username, password = get_credentials()

        # Get impersonated user if needed
        impersonate = input(f"{Fore.GREEN}Do you want to impersonate another user? (y/n): {Style.RESET_ALL}").lower() == 'y'
        impersonated_user = None
        if impersonate:
            impersonated_user = input(f"{Fore.GREEN}Enter email address to impersonate: {Style.RESET_ALL}")

    try:
        # Avant de se connecter au serveur, installer l'intercepteur d'appels EWS
        intercept_ews_calls()
        print(f"{Fore.CYAN}Monitoring individuel des appels EWS activé{Style.RESET_ALL}")
        
        # Connect to account
        print(f"\n{Fore.CYAN}Connecting to Exchange server {server}...{Style.RESET_ALL}")
        
        # Mettre à jour la fonction connect_to_account pour utiliser le serveur spécifié
        def connect_with_server(username, password, impersonated_user=None):
            credentials = Credentials(username, password)
            config = Configuration(server=server, credentials=credentials)
            
            if impersonated_user:
                return Account(
                    primary_smtp_address=impersonated_user,
                    config=config,
                    access_type=DELEGATE,
                    autodiscover=False
                )
            else:
                return Account(
                    primary_smtp_address=username,
                    config=config,
                    access_type=DELEGATE,
                    autodiscover=False
                )
        
        account = connect_with_server(username, password, impersonated_user)
        print(f"{Fore.GREEN}Connected successfully to {Fore.YELLOW}{account.primary_smtp_address}{Style.RESET_ALL}")
        
        if use_unified_interface:
            # Utiliser l'interface unifiée mais SANS démarrer le monitoring automatiquement
            ews_unified_interface.start(start_monitoring=auto_monitor)
            
            # Récupérer la liste des dossiers
            print(f"\n{Fore.CYAN}Listing folders...{Style.RESET_ALL}")
            folders = []
            for i, folder in enumerate(account.root.walk(), 1):
                folder_info = {
                    "index": i,
                    "name": folder.name,
                    "path": folder.absolute,
                    "total_count": folder.total_count,
                    "object": folder
                }
                folders.append(folder_info)
                ews_unified_interface.add_log(f"Found folder: {folder.name} ({folder.total_count} items)")
            
            # Envoyer les dossiers à l'interface
            ews_unified_interface.add_folders(folders)
            
            # Variable pour suivre l'état de l'interface
            current_view = "main"  # "main" ou "stats"
            
            # Afficher les instructions
            print(f"\n{Fore.YELLOW}======= COMMANDES DISPONIBLES ======={Style.RESET_ALL}")
            print(f"{Fore.CYAN}Entrez une commande:{Style.RESET_ALL}")
            print(f"{Fore.GREEN}  1-9{Style.RESET_ALL} : Sélectionner un dossier par son numéro")
            print(f"{Fore.GREEN}  m{Style.RESET_ALL}   : Activer/désactiver le monitoring EWS")
            print(f"{Fore.GREEN}  s{Style.RESET_ALL}   : Afficher les statistiques EWS")
            print(f"{Fore.GREEN}  q{Style.RESET_ALL}   : Quitter le programme")
            
            # Afficher les dossiers
            print(f"\n{Fore.YELLOW}======= DOSSIERS DISPONIBLES ======={Style.RESET_ALL}")
            for i, folder in enumerate(folders, 1):
                print(f"{Fore.WHITE}{i}. {Fore.GREEN}{folder['name']} {Fore.CYAN}({folder['total_count']} items){Style.RESET_ALL}")
            
            # Utiliser un mode d'entrée simple pour Linux qui est plus fiable
            print(f"\n{Fore.YELLOW}Mode d'entrée standard pour Linux activé.{Style.RESET_ALL}")
            
            # Boucle principale pour traiter les commandes de l'interface
            try:
                while True:
                    # Afficher le prompt et attendre l'entrée
                    user_input = input(f"{Fore.GREEN}Commande > {Style.RESET_ALL}").strip().lower()
                    
                    # Aucune entrée, continuer
                    if not user_input:
                        continue
                    
                    # Prendre le premier caractère comme commande
                    key = user_input[0]
                    
                    # Traiter la commande
                    if key == 'q':
                        # Quitter l'application
                        if current_view == "stats":
                            # Si on est en mode stats, revenir au menu
                            ews_stats_window.exit_stats_window()
                            current_view = "main"
                            print(f"\n{Fore.YELLOW}======= DE RETOUR AU MENU PRINCIPAL ======={Style.RESET_ALL}")
                            print(f"{Fore.CYAN}Utilisez 's' pour voir les statistiques ou 'q' pour quitter{Style.RESET_ALL}")
                        else:
                            # Sinon quitter l'application
                            print(f"\n{Fore.YELLOW}Sortie du programme...{Style.RESET_ALL}")
                            break
                    elif key == 'm':
                        # Activer/désactiver le monitoring EWS
                        if ews_unified_interface.monitoring_active:
                            # Désactiver le monitoring
                            ews_unified_interface.stop()
                            ews_unified_interface.start(start_monitoring=False)
                            print(f"{Fore.YELLOW}Monitoring EWS désactivé{Style.RESET_ALL}")
                        else:
                            # Activer le monitoring
                            if ews_unified_interface.start_monitoring():
                                print(f"{Fore.GREEN}Monitoring EWS activé. Les requêtes EWS seront affichées dans une fenêtre séparée.{Style.RESET_ALL}")
                            else:
                                print(f"{Fore.RED}Impossible d'activer le monitoring EWS.{Style.RESET_ALL}")
                    elif key == 's' and current_view == "main":
                        # Afficher les statistiques
                        current_view = "stats"
                        print(f"\n{Fore.YELLOW}======= AFFICHAGE DES STATISTIQUES ======={Style.RESET_ALL}")
                        print(f"{Fore.CYAN}Utilisez 'q' pour revenir au menu principal{Style.RESET_ALL}")
                        ews_stats_window.show_stats_window()
                    elif key.isdigit() and current_view == "main":
                        # Essayer de convertir l'entrée complète en cas de numéro à plusieurs chiffres
                        try:
                            folder_num = int(user_input)
                            if 1 <= folder_num <= len(folders):
                                selected_folder = folders[folder_num-1]
                                print(f"\n{Fore.YELLOW}Traitement du dossier: {Fore.GREEN}{selected_folder['name']}{Style.RESET_ALL}")
                                
                                # Demander confirmation
                                print(f"{Fore.RED}Voulez-vous vraiment vider ce dossier contenant {selected_folder['total_count']} éléments? (o/n){Style.RESET_ALL}")
                                
                                # Lire la confirmation
                                confirm = input(f"{Fore.GREEN}Confirmation > {Style.RESET_ALL}").strip().lower()
                                
                                if confirm == 'o':
                                    # Activer le monitoring si ce n'est pas déjà fait
                                    if not ews_unified_interface.monitoring_active:
                                        print(f"{Fore.YELLOW}Activation du monitoring EWS pour le suivi des opérations...{Style.RESET_ALL}")
                                        ews_unified_interface.start_monitoring()
                                        time.sleep(1)  # Laisser le temps au monitoring de démarrer
                                    
                                    # Vider le dossier
                                    try:
                                        empty_folder(selected_folder["object"])
                                        print(f"\n{Fore.GREEN}Opération terminée.{Style.RESET_ALL}")
                                    except Exception as e:
                                        print(f"\n{Fore.RED}Erreur: {e}{Style.RESET_ALL}")
                                else:
                                    print(f"\n{Fore.YELLOW}Opération annulée.{Style.RESET_ALL}")
                            else:
                                print(f"{Fore.RED}Numéro de dossier invalide. Veuillez entrer un nombre entre 1 et {len(folders)}.{Style.RESET_ALL}")
                        except ValueError:
                            print(f"{Fore.RED}Entrée invalide. Veuillez entrer un nombre pour sélectionner un dossier.{Style.RESET_ALL}")
                    elif key == 'h' or key == '?':
                        # Afficher l'aide
                        print(f"\n{Fore.YELLOW}======= AIDE ======={Style.RESET_ALL}")
                        print(f"{Fore.CYAN}Commandes disponibles:{Style.RESET_ALL}")
                        print(f"{Fore.GREEN}  1-9{Style.RESET_ALL} : Sélectionner un dossier par son numéro")
                        print(f"{Fore.GREEN}  m{Style.RESET_ALL}   : Activer/désactiver le monitoring EWS")
                        print(f"{Fore.GREEN}  s{Style.RESET_ALL}   : Afficher les statistiques EWS")
                        print(f"{Fore.GREEN}  q{Style.RESET_ALL}   : Quitter le programme")
                        print(f"{Fore.GREEN}  h, ?{Style.RESET_ALL} : Afficher cette aide")
                    else:
                        print(f"{Fore.RED}Commande non reconnue. Tapez 'h' pour afficher l'aide.{Style.RESET_ALL}")
                    
                    # Vérifier si l'utilisateur a demandé de quitter les statistiques
                    if ews_stats_window.user_exit_requested:
                        ews_stats_window.user_exit_requested = False
                        current_view = "main"
                        print(f"\n{Fore.YELLOW}======= DE RETOUR AU MENU PRINCIPAL ======={Style.RESET_ALL}")
                        print(f"{Fore.CYAN}Utilisez 's' pour voir les statistiques ou 'q' pour quitter{Style.RESET_ALL}")
                    
                    # Traiter les commandes de l'interface unifiée
                    command = ews_unified_interface.get_command(timeout=0.1)
                    if command:
                        if command["command"] == "quit":
                            break
                        elif command["command"] == "stats":
                            # Passer à la vue statistiques
                            if current_view != "stats":
                                ews_stats_window.show_stats_window()
                                current_view = "stats"
                                print(f"{Fore.YELLOW}Affichage des statistiques EWS. Utilisez 'q' pour revenir au menu principal.{Style.RESET_ALL}")
                        elif command["command"] == "main":
                            # Revenir à la vue principale
                            if current_view != "main":
                                ews_stats_window.exit_stats_window()
                                current_view = "main"
                        elif command["command"] == "process_folder":
                            folder_index = command["folder_index"]
                            if 0 <= folder_index < len(folders):
                                selected_folder = folders[folder_index]

                                # Informer l'interface que le traitement commence
                                ews_unified_interface.data_queue.put({
                                    "type": "processing_start",
                                    "data": {"folder": selected_folder}
                                })

                                ews_unified_interface.add_log(f"Starting processing of folder: {selected_folder['name']}", "INFO")

                                # Lancer le traitement
                                try:
                                    # Empty the folder
                                    empty_folder(selected_folder["object"])

                                    ews_unified_interface.add_log(f"Finished processing folder: {selected_folder['name']}", "INFO")
                                except Exception as e:
                                    ews_unified_interface.add_log(f"Error processing folder: {e}", "ERROR")

                                # Informer l'interface que le traitement est terminé
                                ews_unified_interface.data_queue.put({
                                    "type": "processing_end",
                                    "data": {"success": True}
                                })
                        elif command["command"] == "cancel_processing":
                            ews_unified_interface.add_log("Processing cancelled by user", "WARN")

                    # Petite pause pour éviter de consommer trop de CPU
                    time.sleep(0.1)
            finally:
                pass  # Aucune restauration nécessaire en mode d'entrée standard
        else:
            # Interface classique (code existant)
            # List folders and let user choose
            folders = list_folders(account)
            while True:
                choice = input(f"\n{Fore.GREEN}Enter folder number to empty (or 'q' to quit): {Style.RESET_ALL}")
                if choice.lower() == 'q':
                    break

                try:
                    folder_index = int(choice) - 1
                    if 0 <= folder_index < len(folders):
                        selected_folder = folders[folder_index]

                        # Confirm deletion
                        confirm = input(f"{Fore.RED}Are you sure you want to PERMANENTLY delete all {Fore.YELLOW}{selected_folder.total_count}{Fore.RED} items from '{Fore.YELLOW}{selected_folder.name}{Fore.RED}'? (yes/no): {Style.RESET_ALL}")
                        if confirm.lower() == 'yes':
                            # Empty the folder
                            empty_folder(selected_folder)
                        else:
                            print(f"{Fore.YELLOW}Operation cancelled.{Style.RESET_ALL}")
                    else:
                        print(f"{Fore.RED}Invalid folder number.{Style.RESET_ALL}")
                except ValueError:
                    print(f"{Fore.RED}Please enter a valid number.{Style.RESET_ALL}")
                except Exception as e:
                    print(f"{Fore.RED}Error processing folder: {e}{Style.RESET_ALL}")

    except Exception as e:
        print(f"{Fore.RED}Error: {e}{Style.RESET_ALL}")
    finally:
        # Arrêter tous les processus
        if use_unified_interface:
            ews_unified_interface.stop()
        else:
            ews_logger.stop()
            ews_stats_window.stop()

        # Attendre un moment pour permettre de voir les statistiques
        print(f"\n{Fore.GREEN}Program completed.{Style.RESET_ALL}")


if __name__ == "__main__":
    main() 