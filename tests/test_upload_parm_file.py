import unittest
import sys, os
sys.path.append(os.path.dirname(os.getcwd()))

from requestsloader import RequestsLoader

class TestUploadParmFile(unittest.TestCase):


    def setUp(self):
        self.loader = RequestsLoader()

#Test whether Parm File exixts or not

    def test_parmfile_not_exist(self):
        sys.path.append(os.path.dirname(os.getcwd()))
        with self.assertRaises(FileNotFoundError):
            self.loader.load_workbook('C:/Users/IBM_ADMIN/Desktop/LearnClass/Python/dlr-master/dlr.xlsx')
              
    
#test whether there is any record in Parm File

    def test_no_request_record(self):
        self.loader.load_workbook('C:/Users/IBM_ADMIN/Desktop/LearnClass/Python/dlr-master/dlr.xlsx')
        self.assertTrue(self.loader.get_records())

#test function get_requests()

    def test_function_get_requests(self):
        self.loader.load_workbook('C:/Users/IBM_ADMIN/Desktop/LearnClass/Python/dlr-master/dlr.xlsx')
        self.assertIsNotNone(self.loader.get_requests())


#test function get_requests_str()

    def test_function_get_requests_str(self):
        self.loader.load_workbook('C:/Users/IBM_ADMIN/Desktop/LearnClass/Python/dlr-master/dlr.xlsx')
        l = len(self.loader.get_requests_str())
        self.assertNotEqual(0,l)



    def tearDown(self):
        pass


if __name__ == '__main__':
    #print(sys.path)
    unittest.main()
