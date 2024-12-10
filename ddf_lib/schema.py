from dataclasses import dataclass
from typing import Optional, List
from pathlib import Path
import pandas as pd
from zipfile import ZipFile
from tempfile import TemporaryDirectory


@dataclass
class CDT:
    ids: List[int]
    df: pd.DataFrame
    
    @classmethod
    def parse(cls, cdt_filepath: str) -> Optional["CDT"]:
        try:
            with open(cdt_filepath, "r") as f:
                content = f.read().splitlines()
            
            ids = [id.strip() for id in content[0].split(" #")]
            ids[0] = ids[0][1:]
            
            columns = [c.strip() for c in content[1].split(" #")]
            columns[0] = columns[0][1:]
            
            rows = []
            for row in content[2:]:
                row = [r.strip() for r in row.split(" #")]
                row[0] = row[0][1:]
                rows.append(row)
            
            df = pd.DataFrame(rows, columns=columns)
            return CDT(ids=ids, df=df)
        except FileNotFoundError:
            print(f"File not found: {cdt_filepath}")
            return None
        except PermissionError:
            print(f"Permission denied: {cdt_filepath}")
            return None
        except Exception as e:
            print(f"Unexpected error parsing {cdt_filepath}: {e}")
            return None


@dataclass
class DDF:
    Glazing: Optional[CDT]
    InternalBlinds: Optional[CDT]
    Panes: Optional[CDT]
    WindowGas: Optional[CDT]
    
    @classmethod
    def parse(cls, ddf_filepath: str) -> "DDF":
        ddf_filepath = Path(ddf_filepath)
        
        with TemporaryDirectory() as temp_dir:
            try:
                with ZipFile(ddf_filepath, "r") as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                set = {"Glazing", "InternalBlinds", "Panes", "WindowGas"}
                cdt_dict = {}
                
                for item in set:
                    cdt_filepath = Path(temp_dir) / f"{item}.cdt"
                    cdt = CDT.parse(str(cdt_filepath))
                    cdt_dict[item] = cdt
                
                return DDF(
                    Glazing=cdt_dict["Glazing"],
                    InternalBlinds=cdt_dict["InternalBlinds"],
                    Panes=cdt_dict["Panes"],
                    WindowGas=cdt_dict["WindowGas"],
                )
                
            except Exception as e:
                print(f"Error processing DDF file {ddf_filepath}: {e}")
                return None


ddf_filepath = r"./samples/2A Glazing + Shading.DDF"
ddf = DDF.parse(ddf_filepath)
print(ddf.Glazing.df)