#include "Python.h"
#include <iostream>
#include <string>


int main() {
    std::cout << "Calling python to query database." << std::endl;

    // Initialize the Python interpreter.
    Py_Initialize();

    PyRun_SimpleString("import sys");
    PyRun_SimpleString("sys.path.append(\".\")");

    // Create some Python objects that will later be assigned values.
    PyObject* pName = nullptr;
    PyObject* pModule = nullptr;
    PyObject* pFunc = nullptr;
    PyObject* pArgs = nullptr;
    PyObject* pValue = nullptr;

    // Convert the file name to a Python string.
    pName = PyUnicode_FromString("sqlite3_test");
    // Import the file as a Python module.
    pModule = PyImport_Import(pName);

    if (pModule != nullptr)
    {
        //get function from module
        pFunc = PyObject_GetAttrString(pModule, "query_database");

        if(pFunc && PyCallable_Check(pFunc))
        {
            // Create a Python tuple to hold the arguments to the method.
            pArgs = PyTuple_New(1);
            // Convert string to a Python unicode.
            pValue = PyUnicode_FromString("SELECT * FROM stocks");

            if(!pValue)
            {
                std::cout<<"Cannot convert argument\n"<<std::endl;
                return 1;
            }
            // Set the Python int as the first and second arguments to the method.
            PyTuple_SetItem(pArgs, 0, pValue);
            //PyTuple_SetItem(pArgs, 1, pValue);
            // Call the function with the arguments.
            PyObject* pResult = PyObject_CallObject(pFunc, pArgs);

            // Print a message if calling the method failed.
            if(pResult == nullptr)
            {
                std::cout<<"Calling the add method failed.\n"<<std::endl;
                return 1;
            }
            else
            {
                // Convert the result to a long from a Python object.
                char* result = static_cast<char*>(PyUnicode_DATA(pResult));
                // Print the result.
                std::cout<<"The result is \n"<<result<<std::endl;
            }

        }
    }
    else {
        PyErr_Print();
        std::cout<<"Failed to load the module\n"<<std::endl;
        return 1;
    }

    // Destroy the Python interpreter.
    if (Py_FinalizeEx() < 0) {
        return 120;
    }
    return 0;
}
