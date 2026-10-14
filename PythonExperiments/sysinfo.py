"""
Informatie filteren - systeeminformatie verzamelen, filteren, opslaan en valideren.
 
Werkt op Windows, Linux (bv. Ubuntu-VM) en macOS, en kan gewoon met
'Run Python File' in VS Code gestart worden (geen Jupyter nodig).
"""
 
import csv
import json
import platform
import re
import socket
import subprocess
from pathlib import Path
 
import psutil
 
# Alle bestanden worden naast dit script opgeslagen, ongeacht de werkmap van VS Code.
BASIS_MAP = Path(__file__).resolve().parent
CSV_BESTAND = BASIS_MAP / "systeeminformatie.csv"
XLSX_BESTAND = BASIS_MAP / "systeeminformatie.xlsx"
SAMENGEVOEGD_CSV = BASIS_MAP / "verzamelde_systeeminformatie.csv"
 
OS_NAAM = platform.system()  # "Windows", "Linux" of "Darwin"
 
 
# ----------------------------------------------------------------------
# Hulpfuncties (per besturingssysteem)
# ----------------------------------------------------------------------
def _run(cmd):
    """Voer een commando uit (zonder shell) en geef de tekstuitvoer terug, of None."""
    try:
        out = subprocess.check_output(
            cmd, stderr=subprocess.DEVNULL, timeout=15
        ).decode(errors="ignore").strip()
        return out or None
    except Exception:
        return None
 
 
def _lees_bestand(pad):
    try:
        with open(pad, "r", encoding="utf-8", errors="ignore") as f:
            tekst = f.read().strip()
        return tekst or None
    except Exception:
        return None
 
 
def _powershell(query):
    return _run(["powershell", "-NoProfile", "-Command", query])
 
 
def get_computer_model():
    """Model van de computer (in een VM bv. 'VMware Virtual Platform')."""
    if OS_NAAM == "Windows":
        model = _powershell(
            "(Get-CimInstance -ClassName Win32_ComputerSystem).Model"
        )
    elif OS_NAAM == "Linux":
        model = _lees_bestand("/sys/class/dmi/id/product_name")
    elif OS_NAAM == "Darwin":
        model = _run(["sysctl", "-n", "hw.model"])
    else:
        model = None
    return model or "Onbekend"
 
 
def get_serial_number():
    """Serienummer. Op Linux is dit vaak alleen leesbaar als root."""
    if OS_NAAM == "Windows":
        serial = _powershell("(Get-CimInstance -ClassName Win32_BIOS).SerialNumber")
    elif OS_NAAM == "Linux":
        serial = _lees_bestand("/sys/class/dmi/id/product_serial")
    elif OS_NAAM == "Darwin":
        serial = _run(
            ["bash", "-c",
             "ioreg -l | awk -F'\"' '/IOPlatformSerialNumber/{print $4}'"]
        )
    else:
        serial = None
    return serial or "Onbekend"
 
 
def get_processor_name():
    """Leesbare naam van de processor (platform.processor() is op Linux vaak leeg)."""
    if OS_NAAM == "Linux":
        info = _lees_bestand("/proc/cpuinfo")
        if info:
            for regel in info.splitlines():
                if regel.lower().startswith(("model name", "hardware", "cpu model")):
                    return regel.split(":", 1)[1].strip()
    if OS_NAAM == "Darwin":
        naam = _run(["sysctl", "-n", "machdep.cpu.brand_string"])
        if naam:
            return naam
    return platform.processor() or platform.machine() or "Onbekend"
 
 
# ----------------------------------------------------------------------
# Verzamelen
# ----------------------------------------------------------------------
def collect_system_info():
    """Haalt alle ruwe systeeminformatie op."""
    data = {}
 
    # CPU
    freq = psutil.cpu_freq()
    data["cpu"] = {
        "platform": platform.platform(),
        "processor": get_processor_name(),
        "physical_cores": psutil.cpu_count(logical=False),
        "logical_processors": psutil.cpu_count(logical=True),
        "cpu_freq": freq._asdict() if freq else None,
        "cpu_percent_total": psutil.cpu_percent(interval=0.1),
        "cpu_percent_per_cpu": psutil.cpu_percent(interval=0.1, percpu=True),
    }
 
    # Geheugen
    data["memory"] = psutil.virtual_memory()._asdict()
    data["swap"] = psutil.swap_memory()._asdict()
 
    # Opslag (alle partities)
    storage = {}
    for part in psutil.disk_partitions(all=False):
        try:
            usage = psutil.disk_usage(part.mountpoint)
        except (PermissionError, OSError):
            continue
        storage[part.mountpoint] = {
            "device": part.device,
            "fstype": part.fstype,
            "opts": part.opts,
            "total": usage.total,
            "used": usage.used,
            "free": usage.free,
            "percent": usage.percent,
        }
    data["storage"] = storage
 
    # Netwerk
    net = {}
    addrs = psutil.net_if_addrs()
    stats = psutil.net_if_stats()
    io = psutil.net_io_counters(pernic=True)
    for iface, addr_list in addrs.items():
        net[iface] = {
            "addrs": [
                {**a._asdict(), "family": int(a.family)} for a in addr_list
            ],
            "stats": stats[iface]._asdict() if iface in stats else None,
            "io": io[iface]._asdict() if iface in io else None,
        }
 
    data["network"] = {
        "hostname": socket.gethostname(),
        "fqdn": socket.getfqdn(),
        "interfaces": net,
    }
    return data
 
 
# ----------------------------------------------------------------------
# Filteren
# ----------------------------------------------------------------------
# Adapters die we niet als 'primaire' interface willen (virtueel/loopback/containers)
NEGEER_LIJST = [
    "loopback", "vmware", "virtual", "docker", "vethernet", "mcafee",
    "veth", "br-", "virbr", "vboxnet", "vmnet", "tailscale", "zt",
]
 
 
def _is_negeer(naam):
    n = naam.lower()
    if n == "lo":
        return True
    return any(x in n for x in NEGEER_LIJST)
 
 
def filter_system_info(raw_info):
    """Filtert de ruwe data naar de gevraagde velden."""
    AF_INET = int(socket.AF_INET)
    AF_INET6 = int(socket.AF_INET6)
    AF_LINK = int(psutil.AF_LINK)  # Windows: -1, Linux: 17, macOS: 18
 
    # 1. Primaire netwerkinterface bepalen
    primary_iface = None
    for iface_name, iface_data in raw_info["network"]["interfaces"].items():
        if _is_negeer(iface_name):
            continue
        stats = iface_data.get("stats")
        if stats and stats.get("isup"):
            for addr in iface_data.get("addrs", []):
                if addr.get("family") == AF_INET:
                    ip = addr.get("address", "")
                    if not ip.startswith(("127.", "169.254.")):
                        primary_iface = iface_name
                        break
        if primary_iface:
            break
 
    mac_addr = ipv4_addr = ipv6_addr = ipv6_link_local = None
 
    if primary_iface:
        for addr in raw_info["network"]["interfaces"][primary_iface]["addrs"]:
            fam = addr.get("family")
            ip = addr.get("address", "")
            if fam == AF_LINK:
                mac_addr = ip
            elif fam == AF_INET:
                ipv4_addr = ip
            elif fam == AF_INET6 or ":" in ip:
                ip = ip.split("%")[0]  # Linux/macOS plakken soms '%eth0' achter het adres
                if ip.lower().startswith("fe80"):
                    ipv6_link_local = ip
                else:
                    ipv6_addr = ip
 
    # 2. Opslag: hoofdschijf (C:\ op Windows, / op Linux/macOS)
    storage = raw_info.get("storage", {})
    for kandidaat in ("C:\\", "/"):
        if kandidaat in storage:
            os_drive_key = kandidaat
            break
    else:
        os_drive_key = next(iter(storage), None)
 
    if os_drive_key:
        drive = storage[os_drive_key]
        fstype = drive.get("fstype") or "N/A"
        capacity_gb = round(drive.get("total", 0) / (1024 ** 3), 2)
    else:
        fstype, capacity_gb = "N/A", 0.0
 
    # 3. CPU-frequentie en RAM
    # In een VM geeft psutil soms None of max = 0 terug -> val terug op 'current'
    cpu_freq = raw_info["cpu"].get("cpu_freq") or {}
    max_freq = cpu_freq.get("max") or cpu_freq.get("current") or "N/A"
    ram_gb = round(raw_info["memory"]["total"] / (1024 ** 3), 2)
 
    # 4. Gefilterde dictionary
    return {
        "computernaam": raw_info["network"]["hostname"],
        "computermodel": get_computer_model(),
        "serienummer": get_serial_number(),
        "processor": raw_info["cpu"]["processor"],
        "aantal_fysieke_cores": raw_info["cpu"]["physical_cores"],
        "max_cpu_frequentie_mhz": max_freq,
        "totale_ram_gb": ram_gb,
        "bestandssysteem_opslag": fstype,
        "opslagcapaciteit_gb": capacity_gb,
        "mac_adres": mac_addr,
        "ipv4_adres": ipv4_addr,
        "ipv6_adres": ipv6_addr,
        "ipv6_link_local": ipv6_link_local,
    }
 
 
# ----------------------------------------------------------------------
# Opslaan
# ----------------------------------------------------------------------
def schrijf_csv(gefilterde_data):
    fieldnames = list(gefilterde_data.keys())
    with open(CSV_BESTAND, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        writer.writerow(gefilterde_data)
    print(f"CSV-bestand '{CSV_BESTAND.name}' succesvol aangemaakt.")
 
 
def schrijf_xlsx(gefilterde_data):
    try:
        import openpyxl
    except ImportError:
        print("openpyxl ontbreekt: installeer met 'pip install openpyxl'. "
              "Spreadsheet wordt overgeslagen.")
        return
 
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Systeeminformatie"
    ws.append(list(gefilterde_data.keys()))
    ws.append(list(gefilterde_data.values()))
    wb.save(XLSX_BESTAND)
    print(f"Spreadsheet '{XLSX_BESTAND.name}' succesvol aangemaakt.")
 
 
# ----------------------------------------------------------------------
# Samenvoegen van alle CSV's in de map
# ----------------------------------------------------------------------
def voeg_csv_samen():
    alle_rijen = []
    fieldnames = []
 
    for pad in sorted(BASIS_MAP.glob("*.csv")):
        if pad.name == SAMENGEVOEGD_CSV.name:
            continue
        with open(pad, "r", encoding="utf-8", newline="") as f:
            reader = csv.DictReader(f, delimiter=";")
            for veld in reader.fieldnames or []:
                if veld not in fieldnames:
                    fieldnames.append(veld)
            alle_rijen.extend(reader)
 
    if fieldnames and alle_rijen:
        with open(SAMENGEVOEGD_CSV, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";", restval="")
            writer.writeheader()
            writer.writerows(alle_rijen)
        print(f"Dataset succesvol samengevoegd in '{SAMENGEVOEGD_CSV.name}'. "
              f"Totaal aantal computers: {len(alle_rijen)}")
    else:
        print("Geen CSV-bestanden gevonden om samen te voegen.")
 
 
# ----------------------------------------------------------------------
# Valideren
# ----------------------------------------------------------------------
IPV4_PATROON = r"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$"
MAC_PATROON = r"^([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})$"
 
 
def valideer(csv_pad=SAMENGEVOEGD_CSV):
    print("--- START AUTOMATISCHE VALIDATIE ---\n")
    if not Path(csv_pad).exists():
        print(f"Bestand '{Path(csv_pad).name}' niet gevonden.")
        return
 
    with open(csv_pad, "r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        for i, rij in enumerate(reader, start=1):
            computer = rij.get("computernaam") or f"Computer {i}"
            print(f"Controleren van: {computer}")
 
            # 1. Lege/ontbrekende velden
            for key, value in rij.items():
                if not value or value in ("N/A", "None", "Onbekend"):
                    print(f"  ⚠️ Waarschuwing: Veld '{key}' is leeg of onbekend.")
 
            # 2. RAM en cores
            try:
                if float(rij.get("totale_ram_gb") or 0) <= 0:
                    print("  ❌ Fout: RAM-geheugen moet groter zijn dan 0 GB.")
            except ValueError:
                print("  ❌ Fout: RAM-geheugen is geen geldig getal.")
 
            try:
                if int(rij.get("aantal_fysieke_cores") or 0) < 1:
                    print("  ❌ Fout: Aantal fysieke cores moet minstens 1 zijn.")
            except ValueError:
                print("  ❌ Fout: Aantal cores is geen geldig geheel getal.")
 
            # 3. IPv4
            ipv4 = rij.get("ipv4_adres") or ""
            if not re.match(IPV4_PATROON, ipv4):
                print(f"  ❌ Fout: Ongeldig IPv4-adres formaat: '{ipv4}'")
 
            # 4. MAC
            mac = rij.get("mac_adres") or ""
            if not re.match(MAC_PATROON, mac):
                print(f"  ❌ Fout: Ongeldig MAC-adres formaat: '{mac}'")
 
            print("  ----")
    print("\n--- VALIDATIE VOLTOOID ---")
 
 
# ----------------------------------------------------------------------
def main():
    raw_data = collect_system_info()
    gefilterde_data = filter_system_info(raw_data)
 
    print(json.dumps(gefilterde_data, indent=4, ensure_ascii=False))
    print()
 
    schrijf_csv(gefilterde_data)
    schrijf_xlsx(gefilterde_data)
    voeg_csv_samen()
    print()
    valideer()
 
 
if __name__ == "__main__":
    main()
 
