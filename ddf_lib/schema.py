from dataclasses import dataclass
from typing import Optional, List
from pathlib import Path
import pandas as pd


@dataclass
class CDT:
    ids: List[int]
    df: pd.DataFrame

    @classmethod
    def parse(cls, cdt_filepath: str) -> Optional["CDT"]:
        try:
            with open(cdt_filepath, "r") as f:
                content = f.read().splitlines()

            # Skip IDs. Extract columns from second line.
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
        ddf_dir = ddf_filepath.parent
        folder_name = ddf_filepath.stem

        set = {"Glazing", "InternalBlinds", "Panes", "WindowGas"}
        cdt_dict = {}
        for item in set:
            cdt_filepath = ddf_dir / folder_name / f"{item}.cdt"
            cdt = CDT.parse(cdt_filepath)
            cdt_dict.update({item: cdt})
        
        glazing = cdt_dict["Glazing"]
        internal_blinds = cdt_dict["InternalBlinds"]
        panes = cdt_dict["Panes"]
        window_gas = cdt_dict["WindowGas"]

        return DDF(
            Glazing=glazing,
            InternalBlinds=internal_blinds,
            Panes=panes,
            WindowGas=window_gas,
        )


# Example usage
ddf_filepath = r".\samples\2A Glazing + Shading.DDF"
ddf = DDF.parse(ddf_filepath)

print(ddf.Glazing.df)
