from ddf_lib import DDF

ddf_filepath = r"samples\construction and materials.DDF"
ddf = DDF.open(ddf_filepath)
print(ddf.available_attributes)