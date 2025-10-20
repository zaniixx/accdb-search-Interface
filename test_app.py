import unittest
from app import app, search_databases

class TestApp(unittest.TestCase):
    def setUp(self):
        self.app = app.test_client()

    def test_home_get(self):
        """Test that the home page loads correctly."""
        response = self.app.get('/')
        self.assertEqual(response.status_code, 200)
        data = response.data.decode('utf-8')
        self.assertIn('بحث في قواعد بيانات الناخبين', data)
        self.assertIn('الرقم الشخصي', data)

    def test_home_post_empty_criteria(self):
        """Test POST with empty criteria returns no results."""
        response = self.app.post('/', data={})
        self.assertEqual(response.status_code, 200)
        data = response.data.decode('utf-8')
        self.assertIn('نتائج البحث', data)
        self.assertIn('عدد النتائج: 0', data)

    def test_home_post_with_criteria(self):
        """Test POST with some criteria (should handle even if no results due to DB issues)."""
        data_form = {'perfirst': 'test'}
        response = self.app.post('/', data=data_form)
        self.assertEqual(response.status_code, 200)
        data = response.data.decode('utf-8')
        self.assertIn('نتائج البحث', data)

    def test_search_databases_empty(self):
        """Test search_databases with empty criteria."""
        results = search_databases({})
        self.assertIsInstance(results, list)

    def test_search_databases_with_criteria(self):
        """Test search_databases with criteria."""
        criteria = {'perfirst': 'test'}
        results = search_databases(criteria)
        self.assertIsInstance(results, list)

if __name__ == '__main__':
    unittest.main()