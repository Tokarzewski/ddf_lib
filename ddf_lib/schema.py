from dataclasses import dataclass, fields
from typing import Optional, List
from pathlib import Path
import pandas as pd
from zipfile import ZipFile
from tempfile import TemporaryDirectory

# CDT file format constants
SEPARATOR_HEADER = " #"  # Used for IDs and column headers
SEPARATOR_DATA = "  #"   # Used for data rows
PREFIX = "#"


@dataclass
class CDT:
    ids: List[int]
    df: pd.DataFrame

    @staticmethod
    def _parse_line(line: str, separator: str) -> List[str]:
        """Parse a line by splitting on separator and stripping prefix from first element."""
        parts = [part for part in line.split(separator)]
        parts[0] = parts[0][len(PREFIX):]
        return parts

    @classmethod
    def read(cls, cdt_filepath: str) -> Optional["CDT"]:
        try:
            with open(cdt_filepath, "r") as f:
                content = f.read().splitlines()

            ids = [int(id) for id in cls._parse_line(content[0], SEPARATOR_HEADER)]
            columns = cls._parse_line(content[1], SEPARATOR_HEADER)
            rows = [cls._parse_line(row, SEPARATOR_DATA) for row in content[2:]]

            df = pd.DataFrame(rows, columns=columns)
            return CDT(ids=ids, df=df)
        
        except FileNotFoundError:
            #print(f"File not found: {cdt_filepath}")
            return None
        
        except PermissionError:
            print(f"Permission denied: {cdt_filepath}")
            return None
        
        except Exception as e:
            print(f"Unexpected error parsing {cdt_filepath}: {e}")
            return None
        
    def save(self, cdt_filepath: str) -> None:
        """Save CDT data to a file."""
        with open(cdt_filepath, "w") as f:
            # Write IDs (using header separator)
            ids_line = SEPARATOR_HEADER.join(str(id) for id in self.ids)
            f.write(f"{PREFIX}{ids_line}\n")

            # Write columns (using header separator)
            cols_line = SEPARATOR_HEADER.join(self.df.columns)
            f.write(f"{PREFIX}{cols_line}\n")

            # Write rows (using data separator - efficient method using apply)
            rows_text = self.df.apply(
                lambda row: f"{PREFIX}{SEPARATOR_DATA.join(row.astype(str))}\n",
                axis=1
            )
            f.write(''.join(rows_text))


@dataclass
class DDF:
    Glazing: Optional[CDT]
    InternalBlinds: Optional[CDT]
    Panes: Optional[CDT]
    WindowGas: Optional[CDT]
    Constructions: Optional[CDT]
    Materials: Optional[CDT]
    ActivityTemplates: Optional[CDT]
    ConstructionTemplates: Optional[CDT]
    DHWTemplates: Optional[CDT]
    FacadeTemplates: Optional[CDT]
    GlazingTemplates: Optional[CDT]
    HourlyWeather: Optional[CDT]
    LightingTemplates: Optional[CDT]
    LocalShading: Optional[CDT]
    LocationTemplates: Optional[CDT]
    SBEMHVACSystems: Optional[CDT]
    Schedules: Optional[CDT]

    @property
    def available_attributes(self) -> List[str]:
        """Return list of attribute names that contain data (not None)."""
        return [field.name for field in fields(self)
                if getattr(self, field.name) is not None]

    def has_data(self, attribute: str) -> bool:
        """Check if a specific attribute has data (is not None)."""
        return getattr(self, attribute, None) is not None

    @classmethod
    def read(cls, ddf_filepath: str) -> "DDF":
        ddf_path = Path(ddf_filepath)

        with TemporaryDirectory() as temp_dir:
            try:
                with ZipFile(ddf_path, "r") as zip_ref:
                    zip_ref.extractall(temp_dir)

                temp_path = Path(temp_dir)
                defined_fields = {field.name for field in fields(cls)}
                available_cdts = {f.stem for f in temp_path.glob("*.cdt")}

                unknown = available_cdts - defined_fields
                if unknown:
                    print(f"Unknown CDT files in {ddf_path.name}: {', '.join(unknown)}")

                cdt_dict = {}
                for field_name in defined_fields:
                    cdt_file = temp_path / f"{field_name}.cdt"
                    cdt_dict[field_name] = CDT.read(str(cdt_file))

                return cls(**cdt_dict)
                
            except Exception as e:
                print(f"Error processing DDF file {ddf_path}: {e}")
                return cls(**{field.name: None for field in fields(cls)})

    def save(self, ddf_filepath: str) -> None:
            """Save DDF data to a new file."""
            ddf_path = Path(ddf_filepath)

            with TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)

                # Save each CDT that has data
                for field_name in self.available_attributes:
                    cdt = getattr(self, field_name)
                    if cdt is not None:
                        cdt_file = temp_path / f"{field_name}.cdt"
                        cdt.save(str(cdt_file))

                # Create zip file
                with ZipFile(ddf_path, "w") as zip_ref:
                    for cdt_file in temp_path.glob("*.cdt"):
                        zip_ref.write(cdt_file, cdt_file.name)