"""
CDR Analyzer v3.1 - Law Enforcement Edition (Production Build)
Generates comprehensive forensic analysis matching i9/JAS format
Produces 19+ analysis sheets including halt callers, halt CDR, conference detection
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class CDRAnalyzer:
    """Comprehensive CDR analysis engine - i9/JAS compatible"""

    def __init__(self, cleaned_df, metadata):
        self.df = cleaned_df
        self.metadata = metadata
        self.target = metadata.get('input_value', '').strip("'\"")
        self.results = {}

    def analyze_all(self):
        """Run all analysis modules and return results dict"""
        logger.info("\n" + "=" * 80)
        logger.info("RUNNING COMPREHENSIVE CDR ANALYSIS")
        logger.info("=" * 80)

        modules = [
            ('Smart Report', self._smart_report_placeholder),
            ('FULL CDR - PRINT', self._full_cdr_print),
            ('Full CDR', self._full_cdr),
            ('Contact Summary', self._contact_summary),
            ('Contact Duration Summary', self._contact_duration_summary),
            ('Top Callers', self._top_callers),
            ('Top Called', self._top_called),
            ('Roaming Contact Summary', self._roaming_contact_summary),
            ('IMEI Summary', self._imei_summary),
            ('Roaming Summary', self._roaming_summary),
            ('Location Summary', self._location_summary),
            ('Festival Calls', self._festival_calls),
            ('Day Halt', self._day_halt),
            ('Day Halt Caller', self._day_halt_caller),
            ('Day Halt CDR', self._day_halt_cdr),
            ('Night Halt', self._night_halt),
            ('Night Halt Caller', self._night_halt_caller),
            ('Night Halt CDR', self._night_halt_cdr),
            ('Conference or Call Hold', self._conference_call_hold),
            ('IMSI Summary', self._imsi_summary),
            ('Hourly Activity', self._hourly_activity),
            ('Weekday Analysis', self._weekday_analysis),
            ('Common IMEI', self._common_imei),
            ('Rejected Data', self._rejected_data),
        ]

        for name, func in modules:
            try:
                func()
            except Exception as e:
                logger.warning(f"  WARNING: {name} failed: {e}")
                if name not in self.results:
                    self.results[name] = pd.DataFrame()

        logger.info(f"\n  All analysis complete! {len(self.results)} sheets generated")
        return self.results

    # ==============================================================
    # SAFE HELPERS
    # ==============================================================
    @staticmethod
    def _safe_dt(val):
        """Safely convert to datetime, return None on failure"""
        try:
            if val is None or (hasattr(val, '__class__') and pd.isna(val)):
                return None
            return pd.to_datetime(val)
        except Exception:
            return None

    @staticmethod
    def _safe_strftime(val, fmt='%d-%b-%Y'):
        """Safely format datetime, return '' on failure"""
        try:
            if val is None:
                return ''
            if hasattr(val, 'strftime') and pd.notna(val):
                return val.strftime(fmt)
            return ''
        except Exception:
            return ''

    @staticmethod
    def _safe_int(val, default=0):
        """Safely convert to int"""
        try:
            if val is None or (hasattr(val, '__class__') and pd.isna(val)):
                return default
            return int(float(val))
        except Exception:
            return default

    # ==============================================================
    # HELPER: Build per-contact stats
    # ==============================================================
    def _contact_stats(self, df, sort_by='TOTAL CALLS'):
        """Calculate per-contact call statistics"""
        rows = []
        # Filter out empty/nan OTHER_PARTY
        valid_df = df[df['OTHER_PARTY'].astype(str).str.strip().replace({'': pd.NA, 'nan': pd.NA, 'None': pd.NA, '-': pd.NA}).notna()]

        for contact in valid_df['OTHER_PARTY'].unique():
            contact_str = str(contact).strip()
            if contact_str in ('', 'nan', 'None', '-'):
                continue
            cdf = valid_df[valid_df['OTHER_PARTY'] == contact]
            if len(cdf) == 0:
                continue

            out_call = len(cdf[cdf['CALL_TYPE_LABEL'] == 'OUT-CALL'])
            in_sms = len(cdf[cdf['CALL_TYPE_LABEL'] == 'IN-SMS'])
            in_call = len(cdf[cdf['CALL_TYPE_LABEL'] == 'IN-CALL'])
            fwd_call = len(cdf[cdf['CALL_TYPE_LABEL'] == 'FWD-CALL'])
            bsm = len(cdf[cdf['CALL_TYPE_LABEL'].str.contains('BSM', na=False)])
            out_sms = len(cdf[cdf['CALL_TYPE_LABEL'] == 'OUT-SMS'])
            total = out_call + in_sms + in_call + fwd_call + bsm + out_sms
            dur = int(cdf['DUR_SECONDS'].sum())

            op_raw = cdf['OP_LABEL'].iloc[0] if 'OP_LABEL' in cdf.columns and len(cdf) > 0 else ''
            op = '' if str(op_raw).lower() in ('nan', 'nat', 'none', '', 'service sms') else str(op_raw)
            # Try B_PARTY_OPERATOR/B_PARTY_CIRCLE if OP_LABEL not useful
            if not op and 'B_PARTY_OPERATOR' in cdf.columns:
                bp_op = str(cdf['B_PARTY_OPERATOR'].iloc[0])
                if bp_op.lower() not in ('nan', 'nat', 'none', '', '-'):
                    op = bp_op

            circle = ''
            if 'B_PARTY_CIRCLE' in cdf.columns:
                circ_raw = str(cdf['B_PARTY_CIRCLE'].iloc[0])
                if circ_raw.lower() not in ('nan', 'nat', 'none', '', '-'):
                    circle = circ_raw
            elif 'ROAMING_CIRCLE' in cdf.columns:
                circ_raw = str(cdf['ROAMING_CIRCLE'].iloc[0])
                if circ_raw.lower() not in ('nan', 'nat', 'none', '', '-'):
                    circle = circ_raw

            ctype_raw = cdf['CONTACT_TYPE'].iloc[0] if 'CONTACT_TYPE' in cdf.columns and len(cdf) > 0 else 'Mobile'
            ctype = 'Mobile' if str(ctype_raw).lower() in ('nan', 'nat', 'none', '') else str(ctype_raw)

            try:
                fc = cdf['CALL_DATETIME'].dropna().min() if 'CALL_DATETIME' in cdf.columns else None
                lc = cdf['CALL_DATETIME'].dropna().max() if 'CALL_DATETIME' in cdf.columns else None
                days = cdf['CALL_DATETIME'].dropna().dt.date.nunique() if 'CALL_DATETIME' in cdf.columns else 0
            except Exception:
                fc, lc, days = None, None, 0

            rows.append({
                'CONTACT NUMBER': contact_str,
                'OUT-CALL': out_call or '',
                'IN-SMS': in_sms or '',
                'IN-CALL': in_call or '',
                'FWD-CALL': fwd_call or '',
                'BSM': bsm or '',
                'OUT-SMS': out_sms or '',
                'TOTAL CALLS': total,
                'TOTAL DURATION': dur,
                'TELECOM OPERATOR': op,
                'CIRCLE': circle,
                'TYPE': ctype,
                'FIRST CALL': fc,
                'LAST CALL': lc,
                'DAYS CALLED': days
            })

        result = pd.DataFrame(rows)
        if len(result) > 0:
            result = result.sort_values(sort_by, ascending=False)
        return result

    def _caller_stats(self, df):
        """Compact caller stats (for halt caller sheets)"""
        rows = []
        # Filter out empty/nan OTHER_PARTY
        valid_df = df[df['OTHER_PARTY'].astype(str).str.strip().replace({'': pd.NA, 'nan': pd.NA, 'None': pd.NA, '-': pd.NA}).notna()]

        for contact in valid_df['OTHER_PARTY'].unique():
            contact_str = str(contact).strip()
            if contact_str in ('', 'nan', 'None', '-'):
                continue
            cdf = valid_df[valid_df['OTHER_PARTY'] == contact]
            if len(cdf) == 0:
                continue
            out_call = len(cdf[cdf['CALL_TYPE_LABEL'] == 'OUT-CALL'])
            in_sms = len(cdf[cdf['CALL_TYPE_LABEL'] == 'IN-SMS'])
            in_call = len(cdf[cdf['CALL_TYPE_LABEL'] == 'IN-CALL'])
            fwd_call = len(cdf[cdf['CALL_TYPE_LABEL'] == 'FWD-CALL'])
            bsm = len(cdf[cdf['CALL_TYPE_LABEL'].str.contains('BSM', na=False)])
            out_sms = len(cdf[cdf['CALL_TYPE_LABEL'] == 'OUT-SMS'])
            total = out_call + in_sms + in_call + fwd_call + bsm + out_sms
            dur = int(cdf['DUR_SECONDS'].sum())
            rows.append({
                'OTHER PARTY': contact_str,
                'OUT-CALL': out_call or '',
                'IN-SMS': in_sms or '',
                'IN-CALL': in_call or '',
                'FWD-CALL': fwd_call or '',
                'BSM': bsm or '',
                'OUT-SMS': out_sms or '',
                'TOTAL CALLS': total,
                'TOTAL DURATION': dur,
            })
        result = pd.DataFrame(rows)
        if len(result) > 0:
            result = result.sort_values('TOTAL CALLS', ascending=False)
        return result

    def _build_full_cdr_df(self, source_df):
        """Build full CDR dataframe with standard column names"""
        col_map = {
            'TARGET': 'TARGET', 'OTHER_PARTY': 'OTHER PARTY',
            'TELECOM_OPERATOR': 'TELECOM OPERATOR',
            'OP_LABEL': 'OPERATOR NAME',
            'B_PARTY_CIRCLE': 'B-PARTY CIRCLE',
            'CALL_DATE': 'CALL DATE',
            'CALL_TIME': 'CALL TIME', 'DUR_SECONDS': 'DUR',
            'CALL_TYPE_LABEL': 'CALL TYPE', 'IMEI': 'IMEI',
            'FIRST_CELL_ID': 'FIRST CELL ID', 'FIRST_CELL_ADDRESS': 'FIRST CELL ADDRESS',
            'LAST_CELL_ID': 'LAST CELL ID', 'LAST_CELL_ADDRESS': 'LAST CELL ADDRESS',
            'SMSC_NO': 'SMSC NO', 'SERVICE_TYPE': 'SERVICE TYPE',
            'IMSI': 'IMSI', 'CALL_FWD_NO': 'CALL FWD NO',
            'ROAMING_CIRCLE': 'ROAMING CIRCLE', 'SW_MSC_ID': 'SW/MSC ID',
            'CONNECTION_TYPE': 'TOC', 'LRN_NO': 'LRN NO',
        }
        out = pd.DataFrame()
        for src, dst in col_map.items():
            out[dst] = source_df[src] if src in source_df.columns else ''
        return out

    def _time_filter(self, mode='day'):
        """Filter records by day (07-22:59) or night (23-06:59)"""
        valid = self.df[self.df['CALL_DATETIME'].notna()].copy()
        if mode == 'day':
            return valid[(valid['CALL_DATETIME'].dt.hour >= 7) & (valid['CALL_DATETIME'].dt.hour < 23)]
        else:
            return valid[(valid['CALL_DATETIME'].dt.hour >= 23) | (valid['CALL_DATETIME'].dt.hour < 7)]

    def _location_stats(self, df):
        """Calculate location-level stats"""
        if 'FIRST_CELL_ID' not in df.columns:
            return pd.DataFrame()
        filt = df.copy()
        filt['_CID'] = filt['FIRST_CELL_ID'].astype(str).str.strip()
        filt = filt[~filt['_CID'].isin(['', 'nan', 'None', '-', '---'])]
        if len(filt) == 0:
            return pd.DataFrame()
        stats = filt.groupby('_CID').agg(
            ADDRESS=('FIRST_CELL_ADDRESS', 'first'),
            TOTAL_CALLS=('CALL_DATETIME', 'count'),
            FIRST_CALL=('CALL_DATETIME', 'min'),
            LAST_CALL=('CALL_DATETIME', 'max'),
        ).reset_index().sort_values('TOTAL_CALLS', ascending=False)
        stats.columns = ['CELL ID', 'ADDRESS', 'TOTAL CALLS', 'FIRST CALL', 'LAST CALL']
        return stats

    # ==============================================================
    # ANALYSIS MODULES
    # ==============================================================

    def _smart_report_placeholder(self):
        """Placeholder - Smart Report is built by reporter.py"""
        self.results['Smart Report'] = pd.DataFrame({'_': ['placeholder']})
        logger.info("  Smart Report placeholder created (built by reporter)")

    @staticmethod
    def _clean_val(val):
        """Return empty string for nan/None values"""
        s = str(val).strip()
        if s.lower() in ('nan', 'nat', 'none', '', '-'):
            return ''
        return s

    def _full_cdr_print(self):
        """Compact print view"""
        logger.info("  Generating Full CDR (Print)...")
        rows = []
        for _, r in self.df.iterrows():
            op = self._clean_val(r.get('OP_LABEL', ''))
            # Use B_PARTY_CIRCLE if available for better info
            circle = self._clean_val(r.get('B_PARTY_CIRCLE', ''))
            if not circle:
                circle = self._clean_val(r.get('ROAMING_CIRCLE', ''))
            party = self._clean_val(r.get('OTHER_PARTY', ''))
            if op and op != 'Service SMS':
                party += f"\n{op}"
            elif circle:
                party += f"\n{circle}"

            cell = self._clean_val(r.get('FIRST_CELL_ID', ''))
            addr = self._clean_val(r.get('FIRST_CELL_ADDRESS', ''))
            if addr and cell:
                cell += f"\n{addr}"
            elif addr:
                cell = addr

            rows.append({
                'OTHER PARTY': party,
                'CALL DATE-TIME': r.get('CALL_DATETIME', ''),
                'DUR': r.get('DUR_SECONDS', 0),
                'CALL TYPE': r.get('CALL_TYPE_LABEL', ''),
                'IMEI': self._clean_val(r.get('IMEI', '')),
                'FIRST CELL-ID LOCATION': cell
            })
        self.results['FULL CDR - PRINT'] = pd.DataFrame(rows)
        logger.info(f"    {len(rows)} records")

    def _full_cdr(self):
        """Complete CDR with all columns"""
        logger.info("  Generating Full CDR...")
        self.results['Full CDR'] = self._build_full_cdr_df(self.df)
        logger.info(f"    {len(self.df)} records, {len(self.results['Full CDR'].columns)} columns")

    def _contact_summary(self):
        """Contact summary sorted by frequency"""
        logger.info("  Generating Contact Summary...")
        self.results['Contact Summary'] = self._contact_stats(self.df, 'TOTAL CALLS')
        logger.info(f"    {len(self.results['Contact Summary'])} contacts")

    def _contact_duration_summary(self):
        """Contact summary sorted by duration"""
        logger.info("  Generating Contact Duration Summary...")
        stats = self._contact_stats(self.df, 'TOTAL DURATION')
        self.results['Contact Duration Summary'] = stats[stats['TOTAL DURATION'] > 0]
        logger.info(f"    {len(self.results['Contact Duration Summary'])} contacts with duration")

    def _roaming_contact_summary(self):
        """Per-roaming-circle contact breakdown"""
        logger.info("  Generating Roaming Contact Summary...")
        roam_rows = []
        for circle_name in self.df['ROAMING_CIRCLE'].unique():
            if str(circle_name) in ['', 'nan', 'UNKNOWN']:
                continue
            circle_df = self.df[self.df['ROAMING_CIRCLE'] == circle_name]
            cs = self._contact_stats(circle_df)
            for sr, (_, r) in enumerate(cs.iterrows(), 1):
                roam_rows.append({
                    'SR.': sr,
                    'CONTACT NUMBER': r.get('CONTACT NUMBER', ''),
                    'ROAM CIRCLE': circle_name,
                    'OUT-CALL': r.get('OUT-CALL', ''),
                    'IN-SMS': r.get('IN-SMS', ''),
                    'IN-CALL': r.get('IN-CALL', ''),
                    'BSM': r.get('BSM', ''),
                    'OUT-SMS': r.get('OUT-SMS', ''),
                    'TOTAL CALLS': r.get('TOTAL CALLS', 0),
                    'TOTAL DURATION': r.get('TOTAL DURATION', 0),
                    'TELECOM OPERATOR': r.get('TELECOM OPERATOR', ''),
                    'CIRCLE': r.get('CIRCLE', ''),
                    'TYPE': r.get('TYPE', ''),
                    'FIRST CALL': r.get('FIRST CALL'),
                    'LAST CALL': r.get('LAST CALL'),
                    'DAYS CALLED': r.get('DAYS CALLED', 0),
                })

        self.results['Roaming Contact Summary'] = pd.DataFrame(roam_rows)
        logger.info(f"    {len(roam_rows)} entries")

    def _imei_summary(self):
        """IMEI device summary"""
        logger.info("  Generating IMEI Summary...")
        if 'IMEI' not in self.df.columns:
            self.results['IMEI Summary'] = pd.DataFrame()
            return

        rows = []
        for imei in self.df['IMEI'].unique():
            if str(imei).strip() in ['', 'nan']:
                continue
            idf = self.df[self.df['IMEI'] == imei]
            out_call = len(idf[idf['CALL_TYPE_LABEL'] == 'OUT-CALL'])
            in_sms = len(idf[idf['CALL_TYPE_LABEL'] == 'IN-SMS'])
            in_call = len(idf[idf['CALL_TYPE_LABEL'] == 'IN-CALL'])
            bsm = len(idf[idf['CALL_TYPE_LABEL'].str.contains('BSM', na=False)])
            out_sms = len(idf[idf['CALL_TYPE_LABEL'] == 'OUT-SMS'])
            total = out_call + in_sms + in_call + bsm + out_sms
            try:
                fc = idf['CALL_DATETIME'].dropna().min()
                lc = idf['CALL_DATETIME'].dropna().max()
                days = idf['CALL_DATETIME'].dropna().dt.date.nunique()
            except Exception:
                fc, lc, days = pd.NaT, pd.NaT, 0
            rows.append({
                'IMEI': imei, 'OUT-CALL': out_call or '', 'IN-SMS': in_sms or '',
                'IN-CALL': in_call or '', 'BSM': bsm or '', 'OUT-SMS': out_sms or '',
                'TOTAL CALLS': total, 'FIRST CALL': fc, 'LAST CALL': lc,
                'DAYS CALLED': days, 'HANDSET': ''
            })
        self.results['IMEI Summary'] = pd.DataFrame(rows)
        logger.info(f"    {len(rows)} devices")

    def _roaming_summary(self):
        """Circle transition timeline"""
        logger.info("  Generating Roaming Summary...")
        roam_df = self.df[self.df['IS_ROAMING'] == 'Y'].sort_values('CALL_DATETIME')
        if len(roam_df) == 0:
            self.results['Roaming Summary'] = pd.DataFrame()
            return

        periods = []
        cur, start, prev = None, None, None
        for _, r in roam_df.iterrows():
            c = r['ROAMING_CIRCLE']
            t = r['CALL_DATETIME']
            if c != cur:
                if cur is not None:
                    periods.append({'ROAMING CIRCLE': cur, 'FROM DATE-TIME': start, 'TO DATE-TIME': prev})
                cur, start = c, t
            prev = t
        if cur:
            periods.append({'ROAMING CIRCLE': cur, 'FROM DATE-TIME': start, 'TO DATE-TIME': prev})

        self.results['Roaming Summary'] = pd.DataFrame(periods)
        logger.info(f"    {len(periods)} roaming periods")

    def _location_summary(self):
        """Cell tower usage summary"""
        logger.info("  Generating Location Summary...")
        if 'FIRST_CELL_ID' not in self.df.columns:
            self.results['Location Summary'] = pd.DataFrame()
            return

        filt = self.df.copy()
        filt['_CID'] = filt['FIRST_CELL_ID'].astype(str).str.strip()
        filt = filt[~filt['_CID'].isin(['', 'nan', 'None', '-', '---'])]

        if len(filt) == 0:
            self.results['Location Summary'] = pd.DataFrame()
            logger.info("    0 locations")
            return

        stats = filt.groupby('_CID').agg(
            TOTAL_CALLS=('CALL_DATETIME', 'count'),
            CELL_ADDRESS=('FIRST_CELL_ADDRESS', 'first'),
            FIRST_CALL=('CALL_DATETIME', 'min'),
            LAST_CALL=('CALL_DATETIME', 'max'),
        ).reset_index().sort_values('TOTAL_CALLS', ascending=False)
        stats.columns = ['CELL ID', 'TOTAL CALLS', 'CELL ADDRESS', 'FIRST CALL', 'LAST CALL']
        self.results['Location Summary'] = stats
        logger.info(f"    {len(stats)} locations")

    def _festival_calls(self):
        """Calls on major festival dates"""
        logger.info("  Generating Festival Calls...")
        if 'CALL_DATETIME' not in self.df.columns:
            self.results['Festival Calls'] = pd.DataFrame()
            return

        festivals = {
            'Holi': {'dates': ['2024-03-25', '2025-03-14', '2026-03-04'], 'religion': 'Hindu'},
            'Eid-ul-Fitr': {'dates': ['2024-04-11', '2025-03-31', '2026-03-21'], 'religion': 'Islamic'},
            'Diwali': {'dates': ['2024-11-01', '2024-11-02', '2025-10-20', '2025-10-21', '2026-11-08', '2026-11-09'], 'religion': 'Hindu'},
            'Christmas': {'dates': ['2024-12-25', '2025-12-25', '2026-12-25'], 'religion': 'Christian'},
            'Lohri': {'dates': ['2024-01-13', '2025-01-13', '2026-01-13'], 'religion': 'Sikh'},
            'Maghi - Lohri': {'dates': ['2024-01-14', '2025-01-14', '2026-01-14'], 'religion': 'Sikh'},
            'Guru Nanak Jayanti': {'dates': ['2024-11-15', '2025-11-05', '2026-10-25'], 'religion': 'Sikh'},
            'Durga Puja': {'dates': ['2024-10-10', '2024-10-11', '2024-10-12', '2025-10-01', '2025-10-02'], 'religion': 'Hindu'},
            'Dussehra': {'dates': ['2024-10-12', '2025-10-02', '2026-09-22'], 'religion': 'Hindu'},
            'Raksha Bandhan': {'dates': ['2024-08-19', '2025-08-09', '2026-08-28'], 'religion': 'Hindu'},
            'Independence Day': {'dates': ['2024-08-15', '2025-08-15', '2026-08-15'], 'religion': 'National'},
            'Republic Day': {'dates': ['2024-01-26', '2025-01-26', '2026-01-26'], 'religion': 'National'},
            'Eid-ul-Adha': {'dates': ['2024-06-17', '2025-06-07', '2026-05-27'], 'religion': 'Islamic'},
            'Mahayana New Year': {'dates': ['2024-01-25', '2025-01-29'], 'religion': 'Buddhist'},
            'Magha Puja Day': {'dates': ['2024-02-24', '2025-02-12'], 'religion': 'Buddhist'},
            'Gandhi Jayanti': {'dates': ['2024-10-02', '2025-10-02', '2026-10-02'], 'religion': 'National'},
            'New Year': {'dates': ['2024-01-01', '2025-01-01', '2026-01-01'], 'religion': 'National'},
            'Ganesh Chaturthi': {'dates': ['2024-09-07', '2025-08-27', '2026-09-15'], 'religion': 'Hindu'},
            'Janmashtami': {'dates': ['2024-08-26', '2025-08-16', '2026-09-04'], 'religion': 'Hindu'},
            'Chhath Puja': {'dates': ['2024-11-07', '2024-11-08', '2025-10-26', '2025-10-27'], 'religion': 'Hindu'},
            'Makar Sankranti': {'dates': ['2024-01-14', '2025-01-14', '2026-01-14'], 'religion': 'Hindu'},
            'Ram Navami': {'dates': ['2024-04-17', '2025-04-06', '2026-03-26'], 'religion': 'Hindu'},
            'Mahashivratri': {'dates': ['2024-03-08', '2025-02-26', '2026-02-15'], 'religion': 'Hindu'},
            'Karwa Chauth': {'dates': ['2024-10-20', '2025-10-10', '2026-10-29'], 'religion': 'Hindu'},
        }

        records = []
        for name, info in festivals.items():
            for d in info['dates']:
                date_obj = pd.to_datetime(d).date()
                try:
                    mask = self.df['CALL_DATETIME'].dt.date == date_obj
                except Exception:
                    continue
                fdf = self.df[mask]
                for _, r in fdf.iterrows():
                    records.append({
                        'RELIGION': info['religion'], 'FESTIVAL NAME': name,
                        'TARGET': r.get('TARGET'), 'OTHER PARTY': r.get('OTHER_PARTY'),
                        'TELECOM OPERATOR': r.get('OP_LABEL'), 'CALL DATE': r.get('CALL_DATE'),
                        'CALL TIME': r.get('CALL_TIME'), 'DUR': r.get('DUR_SECONDS'),
                        'CALL TYPE': r.get('CALL_TYPE_LABEL'), 'IMEI': r.get('IMEI'),
                        'FIRST CELL ID': r.get('FIRST_CELL_ID'),
                        'FIRST CELL ADDRESS': r.get('FIRST_CELL_ADDRESS'),
                    })
        self.results['Festival Calls'] = pd.DataFrame(records)
        logger.info(f"    {len(records)} festival records")

    def _day_halt(self):
        """Day halt location summary (07:00-22:59)"""
        logger.info("  Generating Day Halt...")
        filt = self._time_filter('day')
        self.results['Day Halt'] = self._location_stats(filt)
        logger.info(f"    {len(self.results['Day Halt'])} locations")

    def _day_halt_caller(self):
        """Day halt caller summary"""
        logger.info("  Generating Day Halt Caller...")
        filt = self._time_filter('day')
        self.results['Day Halt Caller'] = self._caller_stats(filt)
        logger.info(f"    {len(self.results['Day Halt Caller'])} callers")

    def _day_halt_cdr(self):
        """Day halt full CDR"""
        logger.info("  Generating Day Halt CDR...")
        filt = self._time_filter('day')
        self.results['Day Halt CDR'] = self._build_full_cdr_df(filt)
        logger.info(f"    {len(self.results['Day Halt CDR'])} records")

    def _night_halt(self):
        """Night halt location summary (23:00-06:59)"""
        logger.info("  Generating Night Halt...")
        filt = self._time_filter('night')
        self.results['Night Halt'] = self._location_stats(filt)
        logger.info(f"    {len(self.results['Night Halt'])} locations")

    def _night_halt_caller(self):
        """Night halt caller summary"""
        logger.info("  Generating Night Halt Caller...")
        filt = self._time_filter('night')
        self.results['Night Halt Caller'] = self._caller_stats(filt)
        logger.info(f"    {len(self.results['Night Halt Caller'])} callers")

    def _night_halt_cdr(self):
        """Night halt full CDR"""
        logger.info("  Generating Night Halt CDR...")
        filt = self._time_filter('night')
        self.results['Night Halt CDR'] = self._build_full_cdr_df(filt)
        logger.info(f"    {len(self.results['Night Halt CDR'])} records")

    def _conference_call_hold(self):
        """Detect potential conference/call hold (overlapping calls or calls starting within another call's duration)"""
        logger.info("  Generating Conference/Call Hold...")
        valid = self.df[self.df['CALL_DATETIME'].notna()].copy()
        valid = valid[valid['DUR_SECONDS'] > 0].sort_values('CALL_DATETIME')
        conf_indices = set()

        if len(valid) > 1:
            dt_arr = valid['CALL_DATETIME'].values
            dur_arr = valid['DUR_SECONDS'].values
            idx_arr = valid.index.values

            for i in range(len(valid) - 1):
                t1 = pd.Timestamp(dt_arr[i])
                d1 = int(dur_arr[i])
                end1 = t1 + pd.Timedelta(seconds=d1)

                # Check next few records for overlap
                for j in range(i + 1, min(i + 10, len(valid))):
                    t2 = pd.Timestamp(dt_arr[j])
                    d2 = int(dur_arr[j])

                    if pd.isna(t1) or pd.isna(t2):
                        continue

                    # Case 1: Call 2 starts during Call 1's duration (overlapping)
                    if t2 < end1 and d2 > 0:
                        conf_indices.add(idx_arr[i])
                        conf_indices.add(idx_arr[j])

                    # Case 2: Calls start within 60s of each other and both have duration
                    diff = abs((t2 - t1).total_seconds())
                    if diff <= 60 and d1 > 0 and d2 > 0:
                        conf_indices.add(idx_arr[i])
                        conf_indices.add(idx_arr[j])

                    # Don't check too far ahead
                    if (t2 - t1).total_seconds() > 3600:
                        break

        if conf_indices:
            conf_df = self.df.loc[list(conf_indices)].sort_values('CALL_DATETIME')
            self.results['Conference or Call Hold'] = self._build_full_cdr_df(conf_df)
        else:
            self.results['Conference or Call Hold'] = pd.DataFrame()
        logger.info(f"    {len(conf_indices)} conference records")

    def _imsi_summary(self):
        """IMSI summary"""
        logger.info("  Generating IMSI Summary...")
        if 'IMSI' not in self.df.columns:
            self.results['IMSI Summary'] = pd.DataFrame()
            return

        rows = []
        for imsi in self.df['IMSI'].unique():
            if str(imsi).strip() in ['', 'nan']:
                continue
            idf = self.df[self.df['IMSI'] == imsi]
            out_call = len(idf[idf['CALL_TYPE_LABEL'] == 'OUT-CALL'])
            in_sms = len(idf[idf['CALL_TYPE_LABEL'] == 'IN-SMS'])
            in_call = len(idf[idf['CALL_TYPE_LABEL'] == 'IN-CALL'])
            bsm = len(idf[idf['CALL_TYPE_LABEL'].str.contains('BSM', na=False)])
            out_sms = len(idf[idf['CALL_TYPE_LABEL'] == 'OUT-SMS'])
            total = out_call + in_sms + in_call + bsm + out_sms
            try:
                fc = idf['CALL_DATETIME'].dropna().min()
                lc = idf['CALL_DATETIME'].dropna().max()
                days = idf['CALL_DATETIME'].dropna().dt.date.nunique()
            except Exception:
                fc, lc, days = pd.NaT, pd.NaT, 0
            rows.append({
                'IMSI': imsi, 'OUT-CALL': out_call or '', 'IN-SMS': in_sms or '',
                'IN-CALL': in_call or '', 'BSM': bsm or '', 'OUT-SMS': out_sms or '',
                'TOTAL CALLS': total, 'FIRST CALL': fc, 'LAST CALL': lc, 'DAYS CALLED': days,
            })
        self.results['IMSI Summary'] = pd.DataFrame(rows)
        logger.info(f"    {len(rows)} IMSI numbers")

    def _rejected_data(self):
        """Collect records that couldn't be parsed properly"""
        logger.info("  Generating Rejected Data...")
        # Find records with missing critical fields
        rejected = self.df[
            (self.df['CALL_DATETIME'].isna()) |
            (self.df['OTHER_PARTY'].astype(str).str.strip().isin(['', 'nan', 'None', '-']))
        ]
        if len(rejected) > 0:
            self.results['Rejected Data'] = self._build_full_cdr_df(rejected)
        else:
            self.results['Rejected Data'] = pd.DataFrame()
        logger.info(f"    {len(rejected)} rejected records")

    def _hourly_activity(self):
        """Call activity breakdown by hour of day"""
        logger.info("  Generating Hourly Activity...")
        if 'CALL_DATETIME' not in self.df.columns:
            self.results['Hourly Activity'] = pd.DataFrame()
            return

        valid = self.df[self.df['CALL_DATETIME'].notna()].copy()
        if len(valid) == 0:
            self.results['Hourly Activity'] = pd.DataFrame()
            return

        valid['HOUR'] = valid['CALL_DATETIME'].dt.hour
        rows = []
        for hour in range(24):
            hdf = valid[valid['HOUR'] == hour]
            out_call = len(hdf[hdf['CALL_TYPE_LABEL'] == 'OUT-CALL'])
            in_call = len(hdf[hdf['CALL_TYPE_LABEL'] == 'IN-CALL'])
            out_sms = len(hdf[hdf['CALL_TYPE_LABEL'] == 'OUT-SMS'])
            in_sms = len(hdf[hdf['CALL_TYPE_LABEL'] == 'IN-SMS'])
            fwd_call = len(hdf[hdf['CALL_TYPE_LABEL'] == 'FWD-CALL'])
            bsm = len(hdf[hdf['CALL_TYPE_LABEL'].str.contains('BSM', na=False)])
            total = len(hdf)
            dur = int(hdf['DUR_SECONDS'].sum()) if 'DUR_SECONDS' in hdf.columns else 0

            h_start = f"{hour:02d}:00"
            h_end = f"{hour:02d}:59"
            rows.append({
                'TIME SLOT': f"{h_start} - {h_end}",
                'OUT-CALL': out_call or '',
                'IN-CALL': in_call or '',
                'FWD-CALL': fwd_call or '',
                'OUT-SMS': out_sms or '',
                'IN-SMS': in_sms or '',
                'BSM': bsm or '',
                'TOTAL': total,
                'TOTAL DURATION': dur,
            })
        self.results['Hourly Activity'] = pd.DataFrame(rows)
        logger.info(f"    24 time slots")

    def _weekday_analysis(self):
        """Call activity breakdown by day of week"""
        logger.info("  Generating Weekday Analysis...")
        if 'CALL_DATETIME' not in self.df.columns:
            self.results['Weekday Analysis'] = pd.DataFrame()
            return

        valid = self.df[self.df['CALL_DATETIME'].notna()].copy()
        if len(valid) == 0:
            self.results['Weekday Analysis'] = pd.DataFrame()
            return

        valid['WEEKDAY'] = valid['CALL_DATETIME'].dt.dayofweek  # 0=Mon
        day_names = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        rows = []
        for day_num in range(7):
            ddf = valid[valid['WEEKDAY'] == day_num]
            out_call = len(ddf[ddf['CALL_TYPE_LABEL'] == 'OUT-CALL'])
            in_call = len(ddf[ddf['CALL_TYPE_LABEL'] == 'IN-CALL'])
            out_sms = len(ddf[ddf['CALL_TYPE_LABEL'] == 'OUT-SMS'])
            in_sms = len(ddf[ddf['CALL_TYPE_LABEL'] == 'IN-SMS'])
            fwd_call = len(ddf[ddf['CALL_TYPE_LABEL'] == 'FWD-CALL'])
            total = len(ddf)
            dur = int(ddf['DUR_SECONDS'].sum()) if 'DUR_SECONDS' in ddf.columns else 0
            rows.append({
                'DAY': day_names[day_num],
                'OUT-CALL': out_call or '',
                'IN-CALL': in_call or '',
                'OUT-SMS': out_sms or '',
                'IN-SMS': in_sms or '',
                'FWD-CALL': fwd_call or '',
                'TOTAL': total,
                'TOTAL DURATION': dur,
            })
        self.results['Weekday Analysis'] = pd.DataFrame(rows)
        logger.info(f"    7 days")

    def _common_imei(self):
        """Find contacts that share the same IMEI (device sharing)"""
        logger.info("  Generating Common IMEI...")
        if 'IMEI' not in self.df.columns:
            self.results['Common IMEI'] = pd.DataFrame()
            return

        valid = self.df[self.df['IMEI'].astype(str).str.strip().replace({'': pd.NA, 'nan': pd.NA}).notna()]
        imei_groups = valid.groupby('IMEI')['OTHER_PARTY'].nunique()
        # Find IMEIs used by/with multiple contacts (interesting for forensics)
        multi = imei_groups[imei_groups > 1]
        if len(multi) == 0:
            self.results['Common IMEI'] = pd.DataFrame()
            logger.info("    No shared IMEIs found")
            return

        rows = []
        for imei in multi.index:
            idf = valid[valid['IMEI'] == imei]
            contacts = idf['OTHER_PARTY'].unique()
            try:
                fc = idf['CALL_DATETIME'].dropna().min()
                lc = idf['CALL_DATETIME'].dropna().max()
            except Exception:
                fc, lc = pd.NaT, pd.NaT
            rows.append({
                'IMEI': imei,
                'CONTACTS': ', '.join(str(c) for c in contacts[:10]),
                'TOTAL CONTACTS': len(contacts),
                'TOTAL RECORDS': len(idf),
                'FIRST USE': fc,
                'LAST USE': lc,
            })
        self.results['Common IMEI'] = pd.DataFrame(rows).sort_values('TOTAL CONTACTS', ascending=False)
        logger.info(f"    {len(rows)} shared IMEIs")

    def _top_callers(self):
        """Top callers - contacts who called target the most"""
        logger.info("  Generating Top Callers...")
        incoming = self.df[self.df['CALL_TYPE_LABEL'].isin(['IN-CALL', 'FWD-CALL'])]
        stats = self._contact_stats(incoming, 'TOTAL CALLS')
        mobile_only = stats[stats['TYPE'] == 'Mobile'] if 'TYPE' in stats.columns and len(stats) > 0 else stats
        self.results['Top Callers'] = mobile_only.head(50) if len(mobile_only) > 0 else pd.DataFrame()
        logger.info(f"    {len(self.results['Top Callers'])} callers")

    def _top_called(self):
        """Top called - contacts target called the most"""
        logger.info("  Generating Top Called...")
        outgoing = self.df[self.df['CALL_TYPE_LABEL'] == 'OUT-CALL']
        stats = self._contact_stats(outgoing, 'TOTAL CALLS')
        mobile_only = stats[stats['TYPE'] == 'Mobile'] if 'TYPE' in stats.columns and len(stats) > 0 else stats
        self.results['Top Called'] = mobile_only.head(50) if len(mobile_only) > 0 else pd.DataFrame()
        logger.info(f"    {len(self.results['Top Called'])} called numbers")


if __name__ == "__main__":
    print("CDR Analyzer v3.0 - Run via gui.py")
