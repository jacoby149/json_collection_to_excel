import pandas as pd
from io import BytesIO

# Recursive function to extract nested lists
def extract_nested_lists(data, parent_meta=None, parent_key='root', dfs=None):
    if dfs is None:
        dfs = {}

    # Handle dictionaries
    if isinstance(data, dict):
        for key, value in data.items():
            if isinstance(value, list):
                # Normalize the list and add parent metadata
                df = pd.json_normalize(value)
                if parent_meta:
                    for meta_key, meta_value in parent_meta.items():
                        df[meta_key] = meta_value

                # Store the DataFrame with a unique name
                dfs[f"{parent_key}_{key}"] = df

                # Recurse into each item of the list
                for item in value:
                    extract_nested_lists(item, parent_meta=parent_meta, parent_key=f"{parent_key}_{key}", dfs=dfs)

            elif isinstance(value, dict):
                # Recurse into nested dictionaries
                extract_nested_lists(value, parent_meta=parent_meta, parent_key=f"{parent_key}_{key}", dfs=dfs)

    # Handle lists at the top level
    elif isinstance(data, list):
        for item in data:
            extract_nested_lists(item, parent_meta=parent_meta, parent_key=parent_key, dfs=dfs)

    return dfs

# Function for converting JSON collection to Excel and returning it as a BytesIO object or saving to a file
def json_collection_to_excel(df, filename=None):
    # Extract nested lists into separate DataFrames
    dfs = extract_nested_lists(df)

    # Create a BytesIO object to store the Excel file in memory
    output = BytesIO()

    # Save all DataFrames to a single Excel file within the BytesIO object
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for key, df in dfs.items():
            df.to_excel(writer, sheet_name=key, index=False)

    # Seek to the beginning of the BytesIO object before returning it
    output.seek(0)

    # If filename is provided, save the output to the specified file
    if filename:
        with open(filename, 'wb') as f:
            f.write(output.read())
        print(f"Excel data has been written to {filename}.")
        return None  # Return None if saved to file
    else:
        return output  # Return the BytesIO object if filename is not provided

# Adding method to DataFrame class
pd.DataFrame.json_collection_to_excel = json_collection_to_excel

if __name__ == "__main__":
    # Sample Nested JSON data
    data = [
        {
            "id": 1,
            "name": "Alice",
            "projects": [
                {"title": "AI Research", "year": 2021, "team": [{"member": "John"}, {"member": "Jane"}]},
                {"title": "ML Project", "year": 2022, "team": [{"member": "Tom"}]}
            ]
        },
        {
            "id": 2,
            "name": "Bob",
            "projects": [
                {"title": "Web Development", "year": 2020, "team": [{"member": "Sarah"}]}
            ]
        }
    ]

    # Convert data to DataFrame
    df = pd.DataFrame(data)

    # Call the method on the sample data to generate Excel, either saving to file or returning BytesIO
    filename = 'output.xlsx'  # Set to None if you want to return the BytesIO object
    result = df.json_collection_to_excel(filename=filename)

    if result is None:
        print(f"Excel file saved as {filename}.")
    else:
        print("Excel data has been generated and returned as a BytesIO object.")
