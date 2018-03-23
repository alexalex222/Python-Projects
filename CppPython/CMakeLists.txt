cmake_minimum_required(VERSION 3.9)
project(CppPython)

set(CMAKE_CXX_STANDARD 11)



#Python
set(PYTHON_LIBRARY /Users/kuilinchen/anaconda3/lib/libpython3.6m.dylib)
set(PYTHON_INCLUDE_DIR /Users/kuilinchen/anaconda3/include/python3.6m)
set(PYTHON_EXECUTABLE /Users/kuilinchen/anaconda3/bin/python)
find_package(PythonInterp)
if(PythonInterp_FOUND)
    message(STATUS "PYTHON_DIRS: ${PYTHON_EXECUTABLE}")
endif()
find_package(PythonLibs REQUIRED)
if(PythonLibs_FOUND)
    message(STATUS "PYTHON_INCLUDE_DIRS: ${PYTHON_INCLUDE_DIRS}")
    message(STATUS "PYTHON_LIBRARIES: ${PYTHON_LIBRARIES}")
    message(STATUS "PYTHON_VERSION: ${PYTHON_VERSION}")

    include_directories(${PYTHON_INCLUDE_DIRS})

endif()
add_executable(CppPython main.cpp)
target_link_libraries(CppPython ${PYTHON_LIBRARIES})
