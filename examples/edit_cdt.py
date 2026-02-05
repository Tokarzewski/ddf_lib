from ddf_lib import DDF

# Parse a DDF file
ddf_filepath = r"samples\construction and materials.DDF"
ddf = DDF.read(ddf_filepath)

print(f"Available CDT files: {ddf.available_attributes}")
print()

# Check if Materials data is available
if ddf.has_data("Materials"):
    print("Editing Materials CDT:")
    print(f"  IDs: {ddf.Materials.ids}")
    print(f"  Shape: {ddf.Materials.df.shape}")
    #print(f"  Columns: {list(ddf.Materials.df.columns)}")
    print()

    # Show first few rows before editing
    print("Before editing:")
    print(ddf.Materials.df.head())
    print()

    # Edit a value in the dataframe
    # Change the first row, first data column
    if not ddf.Materials.df.empty:
        second_col = ddf.Materials.df.columns[1]
        original_value = ddf.Materials.df.loc[0, second_col]
        ddf.Materials.df.loc[0, second_col] = "MODIFIED_VALUE"

        # Show first few rows after editing
        print("After editing:")
        print(ddf.Materials.df.head())

        # Save the modified DDF to a new file
        output_filepath = r"samples\modified_construction_and_materials.DDF"
        ddf.save(output_filepath)
        print()
        print(f"Saved modified DDF to: {output_filepath}")
else:
    print("Materials data not available in this DDF file")