import unittest
from unittest.mock import patch
from io import StringIO
import pandas as pd
import main


class TestTravelListGenerator(unittest.TestCase):
    def test_generate_travel_list(self):
        # Test the generate_travel_list function
        destination = "TestDestination"
        activities = ["Activity1", "Activity2"]
        components = ["Component1", "Component2"]
        places = ["Hotel"]
        duration = "Number of days"
        weather = "Cool"
        people = "male"

        expected_result = {
            "Destination": destination,
            "Places": places,
            "Duration": duration,
            "Activities": activities,
            "Components": components,
            "Weather": weather,
            "People": people
        }

        result = main.generate_travel_list(destination, activities, components, places, duration, weather, people)

        self.assertEqual(result, expected_result)

    def test_save_to_excel(self):
        # Test the save_to_excel function
        travel_list = {
            "Destination": "TestDestination",
            "Places": ["Hotel"],
            "Duration": "Number of days",
            "Activities": ["Activity1", "Activity2"],
            "Components": ["Component1", "Component2"],
            "Weather": "Cool",
            "People": "Male"
        }

        filename = "test_save_to_excel.xlsx"

        main.save_to_excel(travel_list, filename)

        # Verify that the file is created and contains the expected data
        with pd.ExcelFile(filename) as xls:
            df = pd.read_excel(xls)
            expected_columns = ["Destination", "Place", "Duration", "Activities", "Components", "Weather", "People"]
            self.assertListEqual(list(df.columns), expected_columns)

            expected_values = ["TestDestination", "Hotel", "Number of days (1 days)", "Activity1\nActivity2",
                               "Component1\nComponent2", ""]
            self.assertListEqual(list(df.iloc[0]), expected_values)

    def test_add_activity(self):
        # Test the add_activity function
        travel_list = {"Activities": ["Activity1"]}

        with patch('builtins.input', return_value='NewActivity'):
            main.add_activity(travel_list)

        self.assertIn("NewActivity", travel_list["Activities"])

    def test_remove_necessary_component(self):
        # Test the remove_activity function
        travel_list = {"Components": ["Component1", "Component2"]}

        with patch('builtins.input', return_value='Component1'):
            main.remove_necessary_component(travel_list)

        self.assertNotIn("Component1", travel_list["Components"])
        self.assertIn("Component2", travel_list["Components"])

    def test_add_necessary_components(self):
        # Test the add_necessary_components function
        travel_list = {"Components": ["Component1"]}

        with patch('builtins.input', side_effect=['NewComponent1', 'NewComponent2', 'C']):
            main.add_necessary_components(travel_list)

        self.assertIn("NewComponent1", travel_list["Components"])
        self.assertIn("NewComponent2", travel_list["Components"])
        self.assertNotIn("C", travel_list["Components"])

    def test_choose_place(self):
        # Test the choose_place function
        travel_list = {"Places": []}

        with patch('builtins.input', return_value='1'):
            main.choose_place(travel_list)

        self.assertIn("Hotel", travel_list["Places"])

    def test_choose_duration(self):
        # Test the choose_duration function
        travel_list = {"Duration": ""}

        with patch('builtins.input', return_value='1'):
            main.choose_duration(travel_list)

        self.assertEqual(travel_list["Duration"], "Number of days (1 days)")

    def test_get_destination(self):
        # Test the get_destination function
        with patch('builtins.input', side_effect=['1', 'TestCountry']):
            destination = main.get_destination()

        self.assertEqual(destination, "TestCountry")

    # Add more test cases for other functions...


if __name__ == '__main__':
    unittest.main()
