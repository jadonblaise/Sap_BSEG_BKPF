import pandas as pd
import numpy as np

class SAPAnalyzer:
    """Class to analyze SAP BSEG and BKPF data and compare it with 
    a provided summary balance sheet"""

    def __init__(self, bseg_path, bkpf_path, summary_path):
        self.bseg_path = bseg_path
        self.bkpf_path = bkpf_path
        self.summary_path = summary_path

        # set Dataframes to None
        self.bseg_df = None
        self.bkpf_df = None
        self.summary_df = None
        self.merged_df = None
        self.comparison_df = None

    def load_data(self):
        """loading BSEG into DF(data frame)"""

        print("Loading BSEG data...")
        try:
            self.bseg_df = pd.read_csv(
                self.bseg_path, 
                sep='|', 
                encoding='latin1',
                on_bad_lines='skip')
        except Exception as e:
            print(f"failed to load BSEG: {e}")
            return
        
        self.bseg_df.columns = self.bseg_df.columns.str.strip().str.lower()
        print("BSEG columns:", self.bseg_df.columns.tolist())
        print('BSEG loaded:', self.bseg_df.shape)

        """loading BKPF into DF"""
        print('loading BKPF data...')
        try:
            self.bkpf_df = pd.read_csv(
                self.bkpf_path, 
                sep='|', 
                encoding='latin1')
        except Exception as e:
            print(f"Failed to load BKPF: {e}")
            return
        
        self.bkpf_df.columns = self.bkpf_df.columns.str.strip().str.lower()
        print("BKPF columns: ", self.bkpf_df.columns.tolist())
        print('BKPF loaded: ', self.bkpf_df.shape)

        """loading summary balance..."""
        print("Loading summary balance sheet...")
        self.summary_df = pd.read_excel(self.summary_path)
        self.summary_df.columns = self.summary_df.columns.str.strip().str.lower()
        print("Summary balance loaded: ", self.summary_df.shape)


    def merge_tables(self):
        """Merges BSEG and BKPF key columns"""
        print("Merging tables...")
        self.merged_df = pd.merge(
            self.bseg_df, 
            self.bkpf_df,
            on='belegnr', # belegnr
            how='inner', # inner
        )
        print('Merged data: ', self.merged_df.shape)

    def validate_data(self):
        """Runs a data check, even if 'bukr' is missing in BSEG """
        print("Validating data completeness...")

        # checks for a missing 'belegnr' in both tables
        for df_name, df in [('BSEG', self.bseg_df), ('BKPF', self.bkpf_df)]:
            if 'belegnr' not in df.columns:
                raise ValueError(f"Missing 'belegnr' column in {df_name} table.")
            if df['belegnr'].isnull().any():
                print(f'Warning: Missing key(s) in {df_name} table - some "belegnr" values are null.')
            
        # Merge only on 'belegnr', because BSEG lacks 'bukr'
        check_df = pd.merge(
            self.bseg_df, 
            self.bkpf_df,
            on='belegnr',
            how='outer', 
            indicator=True
        )
        # identify unmatched records
        missing_headers = check_df[check_df['_merge'] == 'left_only']
        missing_lines = check_df[check_df['_merge'] == 'right_only']
        if not missing_headers.empty:
            print(f"Warning: {len(missing_headers)} BSEG lines without BKPF headers.")
        if not missing_lines.empty:
            print(f"Warning: {len(missing_lines)} BKPF headers without BSEG lines.")
        if missing_headers.empty and missing_lines.empty:
            print("No unmatched records found. Data export is complete.")

    def compare_summary(self):
        """Compare SAP data against summary sheet totals by 'hauptbuch' (GL account)"""
        print('Comparing balances...')

        #Checking required fields
        if 'betrag hauswähr' not in self.merged_df.columns:
            raise KeyError('The column "betrag hauswähr" is missing in merged data')
        if 'hauptbuch' not in self.merged_df.columns or 'hauptbuch' not in self.summary_df.columns:
            raise KeyError('"hauptbuch" is missing either in merged or summary datasheet')
        if 'endsaldo' not in self.summary_df.columns:
            raise KeyError('Missing "endsaldo" column in summary sheet')
        
        # Normalize and clean merged_df amounts
        self.merged_df['betrag hauswähr'] = (
            self.merged_df['betrag hauswähr']
            .astype(str)
            .str.replace('.','', regex=False)
            .str.replace(',','.', regex=False)
            .str.strip()
            .replace('', np.nan)
            .astype(float)
        )
        
        # compute balances per hauptbuch
        balance_computed = (
            self.merged_df.groupby('hauptbuch')['betrag hauswähr']
            .sum()
            .reset_index()
            .rename(columns={'betrag hauswähr': 'computed_balance'})
        )
        # print to verify columns
        print("balance_computed columns:", balance_computed.columns.tolist())

        # clean summary_df
        self.summary_df['endsaldo'] = (
            self.summary_df['endsaldo']
            .astype(str)
            .str.replace('.','', regex=False)
            .str.replace(',','.', regex=False)
            .str.strip()
            .replace('', np.nan)
            .astype(float)
        )
        # print to verify table
        print("summary_df columns:", self.summary_df.columns.tolist())

        # Normalize hauptbuch as 6-character strings with zero padding
        balance_computed['hauptbuch'] = (
        balance_computed['hauptbuch'].astype(str).str.extract(r'(\d+)')[0].str.zfill(6)
        )
        self.summary_df['hauptbuch'] = (
        self.summary_df['hauptbuch'].astype(str).str.extract(r'(\d+)')[0].str.zfill(6)
        )

        # Add row order index before merge
        self.summary_df = self.summary_df.reset_index().rename(columns={'index': 'original_order'})

        # Merge with summary_df as the LEFT dataframe to preserve its order
        self.comparison_df = pd.merge(
            self.summary_df,
            balance_computed,
            on='hauptbuch',
            how='left' # LEFT join to keep summary_df order
        )
    
        # compute difference
        self.comparison_df['difference'] = (
            self.comparison_df['computed_balance'] - self.comparison_df['endsaldo']
        )

        # Sort by original order
        self.comparison_df = self.comparison_df.sort_values(by='original_order').drop(columns=['original_order']) 

        print("Comparison completed:")
        self.comparison_df.rename(columns={'computed_balance': 'endsaldo_computed'}, inplace=True)
        print(self.comparison_df[['hauptbuch', 'endsaldo_computed']])


    
    def export_results(self, output_path):
        """Exports merged data and comparison table to an Excel file."""
        print(f"Exporting results to {output_path}...")
        with pd.ExcelWriter(output_path) as writer:
            self.merged_df.to_excel(writer, sheet_name='AnalyzedData', index=False)
            self.comparison_df[['hauptbuch', 'endsaldo_computed']].to_excel(writer, sheet_name='BalancedComparison', index=False)
            print("Export completed.")

if __name__ == "__main__":
    # instantiate the Analyzer with file paths
    Analyzer = SAPAnalyzer(
        bseg_path="BSEG.csv",
        bkpf_path="BKPF.txt",
        summary_path="Susa_BergerUndCo.xlsx"
    )

    Analyzer.load_data()
    Analyzer.merge_tables()
    Analyzer.validate_data()
    Analyzer.compare_summary()
    #Analyzer.export_results("Abstimmung.xlsx")