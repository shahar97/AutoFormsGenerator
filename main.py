from typing import List
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate

WORD_TEMPLATE = "volunteer_referral_form.docx"
columns_mapping = {
    'לתקופה של': 'period',
    'החל ביום': 'starting',
    'התנדב/ה לעבוד בתפקיד': 'Role',
    'דואר אלקטרוני': 'Email',
    'טלפון נייד': 'number',
    'ישוב': 'City',
    'רחוב, דירה ומספר בית': 'address',
    'שם פרטי ומשפחה': 'full name',
    'מספר זהות': 'ID'

}

columns_to_drop = ['Event Id', 'Id', 'חותמת', 'שם', 'שם וחתימת נותן ההפניה',
                   'למען (לציין זהות הגוף או האדם שלמענו נעשית הפעולה ומקום הפעולה):', 'תאריך',
                   'מלא תאריך סטטוס רישום']


# Get the big list of all full names.
def filter_empty_string(list_of_strings: list):
    # x is a item in list -full name (last name/first name)
    return list(filter(lambda x: len(x) > 1, list_of_strings))


# Replace space holders in word docx with values from Excel .
def replacement(data: pd.DataFrame, output_path: Path):
    # generate doc
    doc = DocxTemplate(WORD_TEMPLATE)
    for index, volunteer in data.iterrows():
        doc.render(volunteer.to_dict())
        doc_name = output_path / f'{volunteer.ID}.docx'
        if doc_name.exists():
            continue
        doc.save(doc_name)


# Split 'Full Name' into first Name ('First') and Last ('Last Name').
def split_full_name(data: pd.DataFrame) -> pd.DataFrame:
    first_and_last = data['full name'].str.split(' ')
    # first_and_last now is big list of small lists - list of all full names of ech volunteer.
    first_and_last = first_and_last.apply(filter_empty_string)
    first_name = first_and_last.apply(lambda x: x[0])
    last_name = first_and_last.apply(lambda x: x[-1])
    data['First'] = first_name
    data['Last'] = last_name

    # Drop the original 'Full Name' column
    data.drop(columns=['full name'], inplace=True)
    return data


# Drop column which are not use.
def drop_invalid_columns(data: pd.DataFrame, columns_to_drop: List[str]) -> pd.DataFrame:
    data = data.drop(columns_to_drop, axis=1)
    return data


# Translate the column name from hebrew to english.
def column_into_english(data: pd.DataFrame) -> pd.DataFrame:
    data.columns = list(map(lambda x: columns_mapping.get(x, ' '), data.columns))
    return data


# Remove the city from address column.
def remove_city_from_address(data: pd.DataFrame) -> pd.DataFrame:
    address = data['address'].str.split(',')
    address = address.apply(lambda x: x[:-1])
    address = address.apply(lambda x: ' '.join(x))
    data['address'] = address
    return data


def remove_lines_with_nan(data: pd.DataFrame) -> pd.DataFrame:
    # Filtering rows with NaN values in any column
    return data.dropna()


def insert_default_value(data: pd.DataFrame, column_name: str) -> pd.DataFrame:
    # Replacing NaN values in column 'B' with a default string 'No Value'
    default_string = ' '
    data[column_name] = data[column_name].fillna(default_string)
    return data


def main(source_path):
    output_path = Path(source_path).parent / 'output'
    output_path.mkdir(exist_ok=True)
    file = Path(source_path)
    data = pd.read_excel(file)

    data = drop_invalid_columns(data, columns_to_drop)
    data = column_into_english(data)
    for column_name in columns_mapping.values():
        if column_name == 'full name':
            continue
        data = insert_default_value(data, column_name)
    data = remove_lines_with_nan(data)
    data = split_full_name(data)
    data = remove_city_from_address(data)
    replacement(data, output_path)


if __name__ == '__main__':
    main()
