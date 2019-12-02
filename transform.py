import pandas as pd


def add_row_with_type_as_test_manual(octane_df, name, description):
    if octane_df.empty:
        row_as_test_manual_list = [1]  # unique_id
    else:
        row_as_test_manual_list = [octane_df.index.max() + 1]  # unique_id
    row_as_test_manual_list.append('test_manual')  # type
    row_as_test_manual_list.append(name)  # name
    row_as_test_manual_list.append('')  # step_type
    row_as_test_manual_list.append(description)  # description
    row_as_test_manual_list.append('')  # step_description
    row_as_test_manual_list.append('')  # test_type
    row_as_test_manual_list.append('')  # product_areas
    row_as_test_manual_list.append('')  # covered_content
    row_as_test_manual_list.append('')  # designer
    row_as_test_manual_list.append('')  # estimated_duration
    row_as_test_manual_list.append('')  # owner
    if octane_df.empty:
        octane_df.loc[1] = row_as_test_manual_list
    else:
        octane_df.loc[octane_df.index.max() + 1] = row_as_test_manual_list


def add_row_with_type_as_step(octane_df, row_type, step_description):
    row_as_step_list = [octane_df.index.max() + 1]  # unique_id
    row_as_step_list.append('step')  # type
    row_as_step_list.append('')  # name
    if row_type == 'Simple':
        row_as_step_list.append('Simple')  # step_type
    if row_type == 'Validation':
        row_as_step_list.append('Validation')  # step_type
    row_as_step_list.append('')  # description
    if row_type == 'Simple':
        row_as_step_list.append(step_description)  # step_description
    if row_type == 'Validation':
        row_as_step_list.append(step_description)  # step_description
    row_as_step_list.append('')  # test_type
    row_as_step_list.append('')  # product_areas
    row_as_step_list.append('')  # covered_content
    row_as_step_list.append('')  # designer
    row_as_step_list.append('')  # estimated_duration
    row_as_step_list.append('')  # owner
    octane_df.loc[octane_df.index.max() + 1] = row_as_step_list


def write_df_to_excel(target_file, octane_df):
    octane_df.to_excel(target_file, sheet_name='manual tests', index=False)
    print("Transformation Completed")


class transform(object):
    SOURCE_FILE = 'ard.xlsx'
    TARGET_FILE = 'octane.xlsx'
    octane_df = pd.DataFrame(
        columns=['unique_id', 'type', 'name', 'step_type', 'description', 'step_description', 'test_type',
                 'product_areas', 'covered_content', 'designer', 'estimated_duration', 'owner'])

    print("Transformation started...please wait")
    ard_df = pd.read_excel(SOURCE_FILE)
    ard_extracted_df = ard_df[['Test Name', 'Description', 'Step Description', 'Expected Result']]

    # Iterate through ARD data frame
    for index, row in ard_extracted_df.iterrows():
        # Find if row['Test Name'] exists in DF; ignore if exists; add to DF with incremented index
        if row['Test Name'] not in octane_df.values:
            # Row with type as 'test_manual'
            add_row_with_type_as_test_manual(octane_df, row['Test Name'], row['Description'])

        # Row with type as 'step' and step_type as 'Simple'
        add_row_with_type_as_step(octane_df, "Simple", row['Step Description'])

        # Row with type as 'step' and step_type as 'Validation'
        add_row_with_type_as_step(octane_df, "Validation", row['Expected Result'])

    write_df_to_excel(TARGET_FILE, octane_df)