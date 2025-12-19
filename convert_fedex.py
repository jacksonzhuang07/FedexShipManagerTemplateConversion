import pandas as pd
import numpy as np
import os

def convert_data():
    source_file = 'test sheet Fed ex .xlsx'
    template_file = 'Hardt_FedEx.csv'
    output_file = 'flashipHardt.csv'
    
    # Read source
    try:
        source_df = pd.read_excel(source_file)
    except Exception as e:
        print(f"Error reading source file: {e}")
        return

    # Read template headers to ensure exact match
    try:
        # We only need the headers
        template_headers = pd.read_csv(template_file, nrows=0).columns.tolist()
    except Exception as e:
        print(f"Error reading template file: {e}")
        return

    # Create output DataFrame with template columns
    output_df = pd.DataFrame(columns=template_headers)
    
    # Process each row in source
    for index, row in source_df.iterrows():
        new_row = {col: '' for col in template_headers}
        
        # Direct Mappings
        new_row['ShipToCompany'] = row.get('to company', '')
        new_row['ShipToContact'] = row.get('to attention', '')
        new_row['ShipToAddress1'] = row.get('to address', '')
        
        # Address 2 logic (Suite/Dept)
        addr2_parts = []
        if pd.notna(row.get('to suite')) and str(row.get('to suite')).strip() != '':
             addr2_parts.append(str(row.get('to suite')))
        if pd.notna(row.get('to department')) and str(row.get('to department')).strip() != '':
             addr2_parts.append(str(row.get('to department')))
        new_row['ShipToAddress2'] = " ".join(addr2_parts)
        
        new_row['ShipToCity'] = row.get('to city', '')
        new_row['ShipToState'] = row.get('to state', '')
        new_row['ShipToCountry'] = row.get('to country', '')
        new_row['ShipToZipCode'] = row.get('to postal code', '')
        new_row['ShipToTel'] = row.get('to phone', '')
        
        # Package Details
        # Format: 10x10x10x25 (L x W x H x Weight) assumed
        pkg_str = str(row.get('packages item 1', ''))
        try:
            dims = pkg_str.lower().split('x')
            if len(dims) >= 4:
                new_row['Length'] = dims[0]
                new_row['Width'] = dims[1]
                new_row['Height'] = dims[2]
                new_row['Weight'] = dims[3]
            else:
                # Fallback or error handling
                print(f"Warning: match format unexpected for row {index}: {pkg_str}")
                # Try to map what we can
                if len(dims) > 0: new_row['Length'] = dims[0]
                if len(dims) > 1: new_row['Width'] = dims[1]
                if len(dims) > 2: new_row['Height'] = dims[2]
        except Exception as e:
             print(f"Error parsing dimensions for row {index}: {e}")

        # Box Count
        new_row['BoxCount'] = 1
        
        # References
        new_row['OrderNum'] = row.get(' reference', '')
        # Assuming PONumber also gets the reference or blank? 
        # Looking at target, PONumber is populated. Let's map reference to PONumber too for visibility.
        new_row['PONumber'] = row.get(' reference', '')
        new_row['Invoice'] = row.get(' reference', '')

        # Static / Template Values (derived from Hardt_FedEx.csv common values)
        new_row['TP_Account'] = '341444179'
        new_row['TP_Bill_Type'] = '3'
        new_row['ServiceType'] = '92' # FedEx Ground
        new_row['PackagingType'] = '1' # Customer Packaging
        new_row['MultiIdentical'] = 'NO'
        
        # Append
        output_df = pd.concat([output_df, pd.DataFrame([new_row])], ignore_index=True)

    # Save
    output_df.to_csv(output_file, index=False)
    print(f"Successfully created {output_file}")

if __name__ == "__main__":
    convert_data()
