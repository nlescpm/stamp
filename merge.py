#!/usr/bin/env python
import argparse
import curses
import json
from Levenshtein import distance
import pathlib
import sys
import pandas as pd
import npyscreen
from mailmerge import MailMerge

parser = argparse.ArgumentParser(description='docx/xlsx Merge tool.')
parser.add_argument('templates', type=pathlib.Path, nargs='+',
                    help='DOCX templates to merge')
parser.add_argument('-x', '--xlsx', required=True, type=pathlib.Path,
                    help="XLSX file")

key_mapping = {}
template_mapping = {}


class MyTestApp(npyscreen.NPSAppManaged):
    """Application wrapper"""

    def onStart(self):
        self.registerForm("MAIN", MainForm())
        self.registerForm("TEMPLATES", TemplatesForm())


class MainForm(npyscreen.SplitForm):
    """This form that will be presented to the user."""

    def create(self):
        """Create the form."""
        self.name = "Define the template"
        self.draw_line_at = 4
        self.add(
                npyscreen.TitleText,
                editable=False,
                name="XLSX file",
                value=args.xlsx.name
                )
        self.add(
                npyscreen.TitleCombo,
                maxlen=1,
                name="Select template by",
                values=list(df.keys()),
                scroll_exit=False,
                )

        for field in fields:
            # find the edit distance to each column
            distances = [distance(field, column) for column in df.columns]

            self.add(
                npyscreen.TitleCombo,
                maxlen=1,
                name=field,
                value=distances.index(min(distances)),
                values=list(df.columns),
                scroll_exit=False,
            )
        self.add_handlers({curses.KEY_DC: self.wipe_value})

    def wipe_value(self, _):
        wipe_widget = self.get_widget(self.editw)
        wipe_widget.set_value(None)

    def afterEditing(self):
        """When 'OK' is pressed."""
        template_by_widget = self.get_widget(1)
        template_by_idx = template_by_widget.get_value()
        if template_by_idx is not None:
            template_by = template_by_widget.get_values()[template_by_idx]

            # find the possible values the template_by key takes
            values_to_map = list(set(df[template_by]))
            values_to_map.sort(key=lambda x: str(x))

            # add a value -> template combobox per value to the next form
            form = self.parentApp.getForm("TEMPLATES")
            for name in values_to_map:
                # find the edit distance to each template
                distances = [distance(name, t.name) for t in args.templates]

                form.add(
                    npyscreen.TitleCombo,
                    maxlen=1,
                    name=str(name),
                    value=distances.index(min(distances)),
                    values=[t.name for t in args.templates],
                    scroll_exit=False,
                )

            # move to the next form
            self.parentApp.setNextForm("TEMPLATES")
        else:
            self.parentApp.setNextForm(None)


class TemplatesForm(npyscreen.Form):
    """This form that will be presented to the user."""

    def create(self):
        """Create the form."""
        self.add_handlers({curses.KEY_DC: self.wipe_value})

    def afterEditing(self):
        """When 'OK' is pressed."""
        for widget in self._widgets__:
            idx = widget.get_value()
            if idx is not None:
                value = widget.get_values()[idx]
                template_mapping[widget.name] = value
            else:
                template_mapping[widget.name] = None

        self.parentApp.setNextForm(None)

    def wipe_value(self, _):
        widget = self.get_widget(self.editw)
        widget.set_value(None)


if __name__ == "__main__":
    args = parser.parse_args()

    df = pd.read_excel(args.xlsx)
    rows, cols = df.shape

    # df.keys() fields in the XLSX to use

    # column 'BCO ' => 'ACCEPT', 'REJECT'
    sfields = set()

    for template in args.templates:
        mm = MailMerge(template)
        mf = mm.get_merge_fields()
        for f in mf:
            sfields.add(f)

    fields = list(sfields)
    fields.sort()

    TA = MyTestApp()
    TA.run()

    for idx, field in enumerate(fields):
        widget = TA.getForm("MAIN").get_widget(idx + 2)

        # make sure we have the right widget
        assert field == widget.name

        v = widget.get_value()
        if v is not None:
            key_mapping[field] = widget.get_values()[v]
        else:
            key_mapping[field] = None

    # TODO: print JSON format of mapping
    print(json.dumps(key_mapping))
    print(json.dumps(template_mapping))

    # doc=  MailMerge('./CIT_2021_Letter-Final-Accept.docx'),

    # # mapping from merge_fields to excel column:
    # mapping = {
    #     'Dossiernummer':     'Dossiernummer',
    #     'Kenmerk_NLeSC':     'Kenmerk NLeSC',
    #     'Title':             'Title',
    #     'Main_applicant':    'Main applicant',
    #     'Affiliation_lvl_1': 'Affiliation (lvl 1)',
    #     'Affiliation_lvl_2': 'Affiliation (lvl 2)',
    #     'Affiliation_lvl_3': 'Affiliation (lvl 3)',
    #     'Address':           'Address',
    #     'Zip_CodeCity':      'Zip Code/City',
    #     'Motivation1':       'Motivation1',
    #     'Motivation2':       'Motivation2',
    #     'Motivation3':       'Motivation3',
    #     'TotaalFTE':         'TotaalFTE',
    #     'WaardeFTE':         'WaardeFTE'
    # }


    # for row in range(rows):
    #     try:
    #         document = templates[df['BCO '][row]]
    #         fields = document.get_merge_fields()
    #         args = {}
    #         for field in fields:
    #             args[field] = df[mapping[field]][row]
    #         document.merge(**args)
    #         document.write(df['Kenmerk NLeSC'][row] + '.docx')
    #         print('Kenmerk:', df['Kenmerk NLeSC'][row], df['BCO '][row])
    #     except:
    #         print('no template for:', df['BCO '][row])
