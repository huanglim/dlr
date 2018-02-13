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
            print('\nParm File  exists.\nPlease go ahead.\n')
        print('\nParm File does not exist! \nPlease check...\n')
    
#test whether there is any record in Parm File

    def test_no_request_record(self):
        self.loader.load_workbook('C:/Users/IBM_ADMIN/Desktop/LearnClass/Python/dlr-master/dlr.xlsx')
        if self.loader.get_records() == None:
            print('\nNo request record in Parm file.\n')
        else:
            print('\nThere are some request records in Parm file.\n')

#test function get_requests()

    def test_function_get_requests(self):
        self.loader.load_workbook('C:/Users/IBM_ADMIN/Desktop/LearnClass/Python/dlr-master/dlr.xlsx')
        if self.loader.get_requests() == None:
            print('No request error!')
        else:
            print('\nRequests listed as below: \n', self.loader.get_requests(), '\n')

#test function get_requests_str()

    def test_function_get_requests_str(self):
        self.loader.load_workbook('C:/Users/IBM_ADMIN/Desktop/LearnClass/Python/dlr-master/dlr.xlsx')
        if self.loader.get_requests_str() == None:
            print('No request error!')
        else:
            print('Requests list as below: \n', self.loader.get_requests_str(), '\n')


    def tearDown(self):
        pass


if __name__ == '__main__':
    #print(sys.path)
    unittest.main()
