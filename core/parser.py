"""
CDR Parser v3.1 - Law Enforcement Edition (Production Build)
Parses CDR files matching the exact format from Indian telecom providers
Supports: Jio, Airtel, Vi, BSNL CSV formats
Bulletproof error handling for field deployment
"""

import pandas as pd
import numpy as np
import re
import os
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class CDRParser:
    """Parse CDR CSV files with metadata header"""

    def __init__(self, csv_path):
        if not csv_path or not os.path.isfile(csv_path):
            raise FileNotFoundError(f"CDR file not found: {csv_path}")
        if os.path.getsize(csv_path) == 0:
            raise ValueError(f"CDR file is empty (0 bytes): {csv_path}")

        self.csv_path = csv_path
        self.raw_df = None
        self.cleaned_df = None
        self._detected_operator = ''

        # Metadata extracted from header
        self.ticket_number = "N/A"
        self.input_value = "Unknown"
        self.date_range = "N/A"
        self.total_records = "N/A"
        self.report_generated_at = "N/A"
        self.msisdn = "N/A"
        self.subscriber_name = "Unknown"
        self.father_husband_name = "N/A"
        self.local_address = "N/A"
        self.circle = "Unknown"
        self.sim_activation_date = "N/A"
        self.port_status = "N/A"

        # Internal
        self.metadata_rows = pd.DataFrame()

    def parse(self):
        """Main parsing function - returns cleaned DataFrame"""
        logger.info("=" * 80)
        logger.info("CDR PARSER v3.1 - LAW ENFORCEMENT EDITION")
        logger.info("=" * 80)

        self._load_csv()

        if self.raw_df is None or len(self.raw_df) == 0:
            raise ValueError("No CDR records found in the file. Check file format.")

        self._extract_metadata()
        self._clean_data()

        if self.cleaned_df is None or len(self.cleaned_df) == 0:
            raise ValueError("All records were invalid after cleaning. Check CDR format.")

        logger.info(f"  DONE: {len(self.cleaned_df)} records ready for analysis")
        return self.cleaned_df

    # ----------------------------------------------------------
    # STEP 1: Load CSV and find where actual CDR data starts
    # ----------------------------------------------------------
    def _load_csv(self):
        logger.info(f"\n  Loading: {self.csv_path}")

        header_row = None
        metadata_lines = []

        # Try multiple encodings
        encoding = self._detect_encoding()

        try:
            with open(self.csv_path, 'r', encoding=encoding) as f:
                for idx, line in enumerate(f):
                    if header_row is None:
                        metadata_lines.append(line.strip())

                    line_lower = line.lower()
                    cdr_keywords = ['target', 'party', 'number', 'call', 'type', 'date', 'time']
                    keyword_count = sum(1 for kw in cdr_keywords if kw in line_lower)

                    if keyword_count >= 5:
                        header_row = idx
                        logger.info(f"  Found CDR header at row {idx}")
                        metadata_lines = metadata_lines[:-1]
                        break

                    if idx > 100:
                        break
        except Exception as e:
            raise ValueError(f"Could not read CSV file: {e}")

        if header_row is None:
            # Try alternate detection: look for "Target No" or "A Party"
            try:
                with open(self.csv_path, 'r', encoding=encoding) as f:
                    for idx, line in enumerate(f):
                        if re.search(r'Target\s*No|A\s*Party|MSISDN', line, re.IGNORECASE):
                            header_row = idx
                            logger.info(f"  Found CDR header (alternate) at row {idx}")
                            break
                        if idx > 200:
                            break
            except Exception:
                pass

        if header_row is None:
            raise ValueError(
                "Could not find CDR header row.\n\n"
                "Expected columns like: Target No, Call Type, B Party No, Date, Time\n"
                "Supported formats: Jio, Airtel, Vi, BSNL CSV exports."
            )

        # Parse metadata lines
        if header_row > 0 and metadata_lines:
            metadata_data = []
            for line in metadata_lines:
                if ',' in line:
                    parts = line.split(',', 1)
                    if len(parts) == 2 and parts[0].strip():
                        metadata_data.append(parts)
            if metadata_data:
                self.metadata_rows = pd.DataFrame(metadata_data, columns=['Field', 'Value'])
            logger.info(f"  Found {len(self.metadata_rows)} metadata rows")

        # Read actual CDR data
        try:
            self.raw_df = pd.read_csv(
                self.csv_path,
                skiprows=header_row,
                encoding=encoding,
                on_bad_lines='skip',
                low_memory=False
            )
        except Exception as e:
            raise ValueError(f"Failed to parse CDR CSV data: {e}")

        # Remove fully empty rows
        if self.raw_df is not None:
            self.raw_df.dropna(how='all', inplace=True)

        rec_count = len(self.raw_df) if self.raw_df is not None else 0
        logger.info(f"  Loaded {rec_count} CDR records")
        if rec_count > 0:
            logger.info(f"  Columns: {list(self.raw_df.columns[:8])}...")

    def _detect_encoding(self):
        """Detect file encoding safely"""
        for enc in ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']:
            try:
                with open(self.csv_path, 'r', encoding=enc) as f:
                    f.read(4096)
                return enc
            except (UnicodeDecodeError, UnicodeError):
                continue
        return 'latin-1'  # Fallback - handles almost anything

    # ----------------------------------------------------------
    # STEP 2: Extract metadata from header rows
    # ----------------------------------------------------------
    def _extract_metadata(self):
        logger.info("\n  Extracting metadata...")

        try:
            self._extract_header_metadata()
        except Exception as e:
            logger.warning(f"  Header metadata extraction error: {e}")

        if len(self.metadata_rows) == 0:
            logger.info("  No key-value metadata rows (using header text extraction)")
            if self.input_value in ('Unknown', 'N/A', ''):
                try:
                    if len(self.raw_df) > 0 and len(self.raw_df.columns) > 0:
                        val = str(self.raw_df.iloc[0, 0]).strip("'\"")
                        # Only use if it looks like a phone number
                        if re.match(r'^\d{10,15}$', val):
                            self.input_value = val
                except Exception:
                    pass
        else:
            field_map = {
                'Ticket Number': 'ticket_number',
                'Input Value': 'input_value',
                'Date Range': 'date_range',
                'Total Records': 'total_records',
                'Report Generated At': 'report_generated_at',
                'MSISDN/IMSI': 'msisdn',
                'Subscriber Name': 'subscriber_name',
                'Father/Husband': 'father_husband_name',
                'Local Address': 'local_address',
                'Circle': 'circle',
                'SIM Activation': 'sim_activation_date',
                'Port in/out': 'port_status',
            }

            for _, row in self.metadata_rows.iterrows():
                try:
                    field = str(row.iloc[0]).strip().rstrip(':')
                    value = str(row.iloc[1]).strip().strip("'\"") if pd.notna(row.iloc[1]) else ''
                    for key, attr in field_map.items():
                        if key in field:
                            setattr(self, attr, value)
                            break
                except Exception:
                    continue

        logger.info(f"  Target: {self.input_value}")
        logger.info(f"  Subscriber: {self.subscriber_name}")
        logger.info(f"  Circle: {self.circle}")
        logger.info(f"  Date Range: {self.date_range}")

    def _extract_header_metadata(self):
        """Extract metadata from free-text header lines like Airtel format"""
        encoding = self._detect_encoding()

        header_lines = []
        try:
            with open(self.csv_path, 'r', encoding=encoding) as f:
                for idx, line in enumerate(f):
                    if idx > 20:
                        break
                    header_lines.append(line.strip())
        except Exception:
            return

        full_header = ' '.join(header_lines[:10])

        # Extract operator name from first line
        for line in header_lines[:5]:
            line_clean = line.strip().rstrip(',')
            if line_clean and not line_clean.startswith(','):
                op_patterns = [
                    (r'BHARTI\s*AIRTEL', 'Airtel'),
                    (r'RELIANCE\s*JIO', 'Jio'),
                    (r'VODAFONE\s*IDEA|Vi\b', 'Vi'),
                    (r'BSNL', 'BSNL'),
                ]
                for pat, op_name in op_patterns:
                    if re.search(pat, line_clean, re.IGNORECASE):
                        self.circle = op_name
                        self._detected_operator = op_name
                        break

        # Extract "Call Details of Mobile No 'XXXXXXXXXX' from 'DD-Mon-YYYY' to 'DD-Mon-YYYY'"
        match = re.search(
            r"Call\s+Details\s+of\s+Mobile\s+No\s*['\"]?(\d+)['\"]?\s*from\s*['\"]?(\d{1,2}-\w{3}-\d{4})['\"]?\s*to\s*['\"]?(\d{1,2}-\w{3}-\d{4})",
            full_header, re.IGNORECASE
        )
        if match:
            self.input_value = match.group(1).strip()
            from_date = match.group(2).strip()
            to_date = match.group(3).strip()
            self.date_range = f"{from_date} to {to_date}"
            self.msisdn = self.input_value
            logger.info(f"  Header: Target={self.input_value}, Range={self.date_range}")

        # Also try IMSI-based: "Call Details of IMSI 'XXXXXXX'"
        if self.input_value in ('Unknown', 'N/A', ''):
            match2 = re.search(
                r"Call\s+Details\s+of\s+(?:IMSI|IMEI|Number)\s*['\"]?(\d+)['\"]?",
                full_header, re.IGNORECASE
            )
            if match2:
                self.input_value = match2.group(1).strip()
                self.msisdn = self.input_value

        # Extract PAN India / circle info
        if re.search(r'PAN\s*India', full_header, re.IGNORECASE):
            if self.circle in ('Unknown', ''):
                self.circle = 'PAN India'
            self._detected_operator = self._detected_operator or self.circle

    # ----------------------------------------------------------
    # STEP 3: Clean and standardize CDR data
    # ----------------------------------------------------------
    def _clean_data(self):
        logger.info("\n  Cleaning data...")

        df = self.raw_df.copy()

        # Remove quotes from all string columns
        for col in df.select_dtypes(include=['object', 'string']).columns:
            try:
                df[col] = df[col].astype(str).str.replace("'", "", regex=False).str.replace('"', '', regex=False).str.strip()
            except Exception:
                pass

        # Comprehensive column mapping (supports Airtel, Jio, Vi, BSNL formats)
        column_mapping = {
            # TARGET / A-Party
            'Target / A Party Number': 'TARGET', 'Target/A Party Number': 'TARGET',
            'A Party Number': 'TARGET', 'Target Number': 'TARGET',
            'Target No': 'TARGET', 'Target no': 'TARGET',
            'MSISDN': 'TARGET', 'A-Party': 'TARGET', 'A Party': 'TARGET', 'A Party No': 'TARGET',

            # Call Type
            'Call Type (In/Out)': 'CALL_TYPE', 'Call Type': 'CALL_TYPE',
            'Type': 'CALL_TYPE', 'In/Out': 'CALL_TYPE',

            # Connection Type
            'Type of Connection': 'CONNECTION_TYPE', 'TOC': 'CONNECTION_TYPE',

            # B-Party / Other Party
            'B Party Number': 'OTHER_PARTY', 'B-Party Number': 'OTHER_PARTY',
            'B Party No': 'OTHER_PARTY', 'B Party': 'OTHER_PARTY',
            'Called Number': 'OTHER_PARTY', 'Other Party': 'OTHER_PARTY', 'Other No': 'OTHER_PARTY',

            # LRN
            'LRN Called No': 'LRN_NO', 'LRN No': 'LRN_NO',

            # Telecom Operator
            'LRN Operator Name with LSA': 'TELECOM_OPERATOR', 'LRN TSP-LSA': 'TELECOM_OPERATOR',
            'Operator': 'TELECOM_OPERATOR', 'Telecom Operator': 'TELECOM_OPERATOR', 'TSP Name': 'TELECOM_OPERATOR',

            # Date / Time
            'Call Date': 'CALL_DATE', 'Date': 'CALL_DATE',
            'Call Time': 'CALL_TIME', 'Time': 'CALL_TIME',

            # Duration
            'Call Duration': 'DUR', 'Duration': 'DUR', 'Dur(s)': 'DUR', 'Dur (s)': 'DUR',
            'Dur': 'DUR', 'Duration(s)': 'DUR', 'Duration (Sec)': 'DUR',

            # Cell / Location
            'First BTS Location-Address': 'FIRST_CELL_ADDRESS', 'First CGI Lat/Long': 'FIRST_CELL_ADDRESS',
            'First Cell Address': 'FIRST_CELL_ADDRESS', 'First Cell ID': 'FIRST_CELL_ID',
            'First CGI': 'FIRST_CELL_ID', 'Cell ID': 'FIRST_CELL_ID', 'Cell Id': 'FIRST_CELL_ID',
            'Last BTS Location-Address': 'LAST_CELL_ADDRESS', 'Last CGI Lat/Long': 'LAST_CELL_ADDRESS',
            'Last Cell Address': 'LAST_CELL_ADDRESS', 'Last Cell ID': 'LAST_CELL_ID', 'Last CGI': 'LAST_CELL_ID',

            # SMSC / Service
            'SMS Center Number': 'SMSC_NO', 'SMSC No': 'SMSC_NO', 'SMSC Number': 'SMSC_NO',
            'Service Type (Voice/SMS)': 'SERVICE_TYPE', 'Service Type': 'SERVICE_TYPE',

            # IMEI / IMSI
            'IMEI': 'IMEI', 'IMSI': 'IMSI',

            # Call Forwarding
            'Call Forwarding Number': 'CALL_FWD_NO', 'Call Fow No': 'CALL_FWD_NO',
            'Call Forward No': 'CALL_FWD_NO', 'CF Number': 'CALL_FWD_NO',

            # Roaming
            'Roaming Circle Name': 'ROAMING_CIRCLE', 'Roaming Circle': 'ROAMING_CIRCLE',
            'Roam Nw': 'ROAMING_CIRCLE', 'Roam Circle': 'ROAMING_CIRCLE',

            # Switch / MSC
            'Switch ID / MSC ID': 'SW_MSC_ID', 'SW & MSC ID': 'SW_MSC_ID', 'MSC ID': 'SW_MSC_ID',

            # Trunk Groups
            'In TG': 'IN_TG', 'IN TG': 'IN_TG', 'Out TG': 'OUT_TG', 'OUT TG': 'OUT_TG',

            # VoWiFi
            'Vowifi First UE IP': 'VOWIFI_FIRST_IP', 'Vowifi Last UE IP': 'VOWIFI_LAST_IP',
        }
        df.rename(columns=column_mapping, inplace=True)

        # Ensure mandatory columns exist
        mandatory = {
            'TARGET': '', 'CALL_TYPE': 'OUT', 'OTHER_PARTY': '',
            'CALL_DATE': '', 'CALL_TIME': '00:00:00', 'DUR': '0',
            'SERVICE_TYPE': 'Voice', 'FIRST_CELL_ID': '', 'FIRST_CELL_ADDRESS': '',
            'LAST_CELL_ID': '', 'LAST_CELL_ADDRESS': '', 'IMEI': '', 'IMSI': '',
            'TELECOM_OPERATOR': '', 'ROAMING_CIRCLE': '', 'CONNECTION_TYPE': 'PREPAID',
            'LRN_NO': '', 'SMSC_NO': '', 'CALL_FWD_NO': '', 'SW_MSC_ID': '',
            'IN_TG': '', 'OUT_TG': ''
        }
        for col, default in mandatory.items():
            if col not in df.columns:
                df[col] = default

        # ---- AIRTEL-SPECIFIC: Split combined lat/long fields ----
        for prefix in ['FIRST', 'LAST']:
            addr_col = f'{prefix}_CELL_ADDRESS'
            lat_col = f'{prefix}_CELL_LAT'
            long_col = f'{prefix}_CELL_LONG'
            try:
                if addr_col in df.columns:
                    addr_vals = df[addr_col].astype(str)
                    is_latlong = addr_vals.str.match(r'^-?\d+\.?\d*/\s*-?\d+\.?\d*$', na=False)
                    if is_latlong.any():
                        split = addr_vals.str.split('/', expand=True, n=1)
                        df[lat_col] = pd.to_numeric(split[0].where(is_latlong), errors='coerce')
                        df[long_col] = pd.to_numeric(split[1].where(is_latlong), errors='coerce') if 1 in split.columns else np.nan
                        df.loc[is_latlong, addr_col] = (
                            df.loc[is_latlong, lat_col].astype(str) + ', ' +
                            df.loc[is_latlong, long_col].astype(str)
                        )
                        logger.info(f"  Airtel lat/long split: {is_latlong.sum()} {prefix} records")
            except Exception as e:
                logger.warning(f"  Lat/long split error ({prefix}): {e}")
            if lat_col not in df.columns:
                df[lat_col] = np.nan
            if long_col not in df.columns:
                df[long_col] = np.nan

        # ---- AIRTEL-SPECIFIC: Roaming circle handling ----
        try:
            if 'ROAMING_CIRCLE' in df.columns:
                roam_vals = df['ROAMING_CIRCLE'].astype(str).str.strip()
                has_space = roam_vals.str.contains(' ', na=False)
                if has_space.any():
                    df['_ROAM_OP'] = roam_vals.apply(
                        lambda x: CDRParser._decode_operator(x.replace(' ', '-'))[0] if ' ' in str(x) else ''
                    )
                    df['_ROAM_CIRCLE'] = roam_vals.apply(
                        lambda x: CDRParser._decode_operator(x.replace(' ', '-'))[1] if ' ' in str(x) else ''
                    )
        except Exception as e:
            logger.warning(f"  Roaming circle parsing error: {e}")

        # Parse duration to seconds (safe)
        df['DUR_SECONDS'] = df['DUR'].apply(self._parse_duration)

        # Parse datetime (safe)
        df['CALL_DATETIME'] = df.apply(
            lambda r: self._parse_datetime(r.get('CALL_DATE', ''), r.get('CALL_TIME', '')), axis=1
        )
        valid_dt = df['CALL_DATETIME'].notna().sum()
        logger.info(f"  Parsed {valid_dt}/{len(df)} datetime values")

        # Standardize call types
        df['CALL_TYPE'] = df['CALL_TYPE'].astype(str).str.upper().str.strip()
        call_type_map = {
            'INCOMING': 'IN', 'OUTGOING': 'OUT',
            'IN CALL': 'IN', 'OUT CALL': 'OUT', 'NAN': 'OUT',
            'MOC': 'OUT', 'MTC': 'IN',
        }
        df['CALL_TYPE'] = df['CALL_TYPE'].replace(call_type_map)

        # Derive labels (each wrapped in try/except)
        df['CALL_TYPE_LABEL'] = df.apply(self._call_type_label, axis=1)
        df['OP_LABEL'] = df.apply(self._operator_label, axis=1)
        df['CONTACT_TYPE'] = df['OTHER_PARTY'].apply(self._contact_type)

        # Roaming flag - smarter detection
        home_circle_code = ''
        try:
            if hasattr(self, 'circle') and self.circle:
                for code, name in CDRParser.CIRCLE_MAP.items():
                    if name.upper() == self.circle.upper():
                        home_circle_code = code
                        break
        except Exception:
            pass

        def _detect_roaming(row):
            try:
                rc = str(row.get('ROAMING_CIRCLE', '')).strip()
                if rc in ('', '-', 'nan', 'UNKNOWN', 'None'):
                    return 'N'
                parts = rc.replace('-', ' ').split()
                if len(parts) >= 2:
                    circle_part = parts[-1].upper()
                    if home_circle_code and circle_part != home_circle_code.upper():
                        return 'Y'
                    return 'N'
                return 'Y' if rc not in ('', '-') else 'N'
            except Exception:
                return 'N'

        df['IS_ROAMING'] = df.apply(_detect_roaming, axis=1)

        # Derive B-Party operator/circle
        df['B_PARTY_OPERATOR'] = df['TELECOM_OPERATOR'].apply(
            lambda x: CDRParser._safe_decode_op(x, 0)
        )
        df['B_PARTY_CIRCLE'] = df['TELECOM_OPERATOR'].apply(
            lambda x: CDRParser._safe_decode_op(x, 1)
        )

        # Remove duplicates and empty rows
        before = len(df)
        try:
            df.drop_duplicates(inplace=True)
            df = df.dropna(how='all')
        except Exception:
            pass
        removed = before - len(df)
        if removed > 0:
            logger.info(f"  Removed {removed} duplicate/empty records")

        # Remove disclaimer/system rows
        try:
            mask = df['TARGET'].astype(str).str.contains(
                'Disclaimer|disclaimer|system generated|NOTE:', na=False, regex=True
            )
            if mask.any():
                df = df[~mask]
                logger.info(f"  Removed {mask.sum()} system/disclaimer rows")
        except Exception:
            pass

        self.cleaned_df = df
        logger.info(f"  Final: {len(df)} records, {len(df.columns)} columns")

    # ----------------------------------------------------------
    # Helper methods (ALL safe - never crash)
    # ----------------------------------------------------------
    @staticmethod
    def _safe_decode_op(x, idx):
        """Safely decode operator code, return op_name (idx=0) or circle_name (idx=1)"""
        try:
            if pd.isna(x):
                return ''
            result = CDRParser._decode_operator(x)
            return result[idx] if result else ''
        except Exception:
            return ''

    @staticmethod
    def _parse_duration(dur_str):
        try:
            if pd.isna(dur_str) or str(dur_str).strip() in ['', 'nan', '-', 'None']:
                return 0
            s = str(dur_str).strip()
            if s.replace('.', '', 1).isdigit():
                return int(float(s))
            if ':' in s:
                parts = list(map(int, s.split(':')))
                if len(parts) == 3:
                    return parts[0] * 3600 + parts[1] * 60 + parts[2]
                if len(parts) == 2:
                    return parts[0] * 60 + parts[1]
            return int(float(s))
        except Exception:
            return 0

    @staticmethod
    def _parse_datetime(date_str, time_str):
        try:
            ds = str(date_str).strip().strip("'\"")
            ts = str(time_str).strip().strip("'\"")
            if ds in ['', 'nan', 'None', '-'] or ts in ['', 'nan', 'None', '-']:
                return pd.NaT
            dts = f"{ds} {ts}"
            for fmt in [
                '%d/%m/%Y %H:%M:%S', '%m/%d/%Y %H:%M:%S',
                '%d-%m-%Y %H:%M:%S', '%Y-%m-%d %H:%M:%S',
                '%d/%m/%Y %H:%M', '%d-%m-%Y %H:%M',
                '%d-%b-%Y %H:%M:%S', '%d-%b-%Y %H:%M',
                '%Y/%m/%d %H:%M:%S', '%Y/%m/%d %H:%M',
                '%d.%m.%Y %H:%M:%S', '%d.%m.%Y %H:%M',
            ]:
                try:
                    return pd.to_datetime(dts, format=fmt)
                except Exception:
                    continue
            return pd.to_datetime(dts, errors='coerce')
        except Exception:
            return pd.NaT

    @staticmethod
    def _call_type_label(row):
        try:
            ct = str(row.get('CALL_TYPE', '')).upper().strip()
            st = str(row.get('SERVICE_TYPE', '')).upper().strip()

            if ct == 'SMT':
                return 'IN-SMS'
            if ct == 'SMO':
                return 'OUT-SMS'
            if ct == 'FOW':
                return 'FWD-CALL'

            if 'A2P' in st or 'SMS' in st:
                if ct in ('IN', 'MTC', 'SMT'):
                    return 'IN-SMS'
                elif ct in ('OUT', 'MOC', 'SMO'):
                    return 'OUT-SMS'
                return 'IN-SMS'
            if 'BSM' in st:
                return 'BSM'

            if ct in ('IN', 'MTC'):
                return 'IN-CALL'
            if ct in ('OUT', 'MOC'):
                return 'OUT-CALL'
            if ct in ('FOW', 'CF', 'FORWARD'):
                return 'FWD-CALL'

            return f"{ct}-CALL" if ct and ct != 'NAN' else "CALL"
        except Exception:
            return 'CALL'

    # Operator code to name mapping for Indian telecom
    OPERATOR_MAP = {
        'AIR': 'Airtel', 'AIRTEL': 'Airtel', 'AT': 'Airtel',
        'RJIL': 'Jio', 'JIO': 'Jio', 'RELJIO': 'Jio', 'RJ': 'Jio',
        'VODA': 'Vodafone', 'IDEA': 'Idea', 'VI': 'Vi', 'VIL': 'Vi', 'VF': 'Vodafone',
        'BSNL': 'BSNL', 'MTNL': 'MTNL',
    }

    # Circle code to name mapping
    CIRCLE_MAP = {
        'DL': 'Delhi', 'MH': 'Maharashtra', 'GJ': 'Gujarat', 'KA': 'Karnataka',
        'TN': 'Tamil Nadu', 'AP': 'Andhra Pradesh', 'TS': 'Telangana', 'KL': 'Kerala',
        'PB': 'Punjab', 'HR': 'Haryana', 'UP': 'UP East', 'UE': 'UP East', 'UW': 'UP West',
        'RJ': 'Rajasthan', 'MP': 'Madhya Pradesh', 'WB': 'West Bengal',
        'BH': 'Bihar', 'BR': 'Bihar', 'OR': 'Odisha', 'OD': 'Odisha',
        'AS': 'Assam', 'NE': 'North East', 'JK': 'J&K', 'HP': 'Himachal Pradesh',
        'KO': 'Kolkata', 'MU': 'Mumbai', 'CH': 'Chennai', 'CG': 'Chhattisgarh',
        'JH': 'Jharkhand', 'PUN': 'Punjab',
    }

    @classmethod
    def _decode_operator(cls, code):
        """Decode telecom operator code like 'AIR-DL' -> ('Airtel', 'Delhi')"""
        try:
            if not code or str(code).strip() in ('', 'nan', '-', 'None'):
                return '', ''
            code = str(code).strip()
            parts = code.split('-')
            op_code = parts[0].upper() if parts else ''
            circle_code = parts[1].upper() if len(parts) > 1 else ''
            op_name = cls.OPERATOR_MAP.get(op_code, code)
            circle_name = cls.CIRCLE_MAP.get(circle_code, circle_code)
            return op_name, circle_name
        except Exception:
            return str(code), ''

    @staticmethod
    def _operator_label(row):
        try:
            ct = str(row.get('CALL_TYPE', '')).upper().strip()
            st = str(row.get('SERVICE_TYPE', '')).upper()
            op = str(row.get('TELECOM_OPERATOR', '')).strip()
            other = str(row.get('OTHER_PARTY', '')).strip()

            if ct in ('SMT', 'SMO') or 'SMS' in st:
                if any(c.isalpha() for c in other):
                    return 'Service SMS'

            if op in ('', 'nan', 'UNKNOWN', '-', 'None'):
                if not any(c.isalpha() for c in other) and len(other) >= 10:
                    roam = str(row.get('ROAMING_CIRCLE', '')).strip()
                    if roam and roam not in ('', 'nan', '-', 'None'):
                        op_name, circle_name = CDRParser._decode_operator(roam.replace(' ', '-'))
                        if circle_name:
                            return f"{op_name} ({circle_name})"
                        return op_name
                if ct in ('SMT', 'SMO') or 'SMS' in st:
                    return 'Service SMS'
                return ''

            op_name, circle_name = CDRParser._decode_operator(op)
            if circle_name:
                return f"{op_name} ({circle_name})"
            return op_name
        except Exception:
            return ''

    @staticmethod
    def _contact_type(party):
        try:
            p = str(party).strip()
            if p in ('', 'nan', 'None', '-'):
                return 'Unknown'
            if p.startswith('00') or (p.startswith('+') and len(p) > 13):
                return 'ISD'
            if any(c.isalpha() for c in p):
                return 'Service SMS'
            digits_only = ''.join(c for c in p if c.isdigit())
            if len(digits_only) <= 6:
                return 'Short Code'
            return 'Mobile'
        except Exception:
            return 'Unknown'

    def get_metadata_dict(self):
        """Return all metadata as a dictionary for other modules"""
        return {
            'ticket_number': self.ticket_number,
            'input_value': self.input_value,
            'date_range': self.date_range,
            'total_records': self.total_records,
            'subscriber_name': self.subscriber_name,
            'father_husband_name': self.father_husband_name,
            'local_address': self.local_address,
            'circle': self.circle,
            'operator': getattr(self, '_detected_operator', self.circle),
            'sim_activation_date': self.sim_activation_date,
            'msisdn': self.msisdn,
        }


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        parser = CDRParser(sys.argv[1])
        df = parser.parse()
        print(f"\nTarget: {parser.input_value}")
        print(f"Subscriber: {parser.subscriber_name}")
        print(f"Shape: {df.shape}")
        print(f"\nFirst 5 records:\n{df.head().to_string()}")
