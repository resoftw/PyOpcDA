#include <Python.h>
#include "OPCClient.h"

/*
 * Implements an example function.
 */
PyDoc_STRVAR(OpcDA_init_doc, "Init(domain,username,password)\
\
Initialize Security");
PyDoc_STRVAR(OpcDA_connect_doc, "Connect(host,progid)\
\
Connect to OPC Server");

PyDoc_STRVAR(OpcDA_disconnect_doc, "Disconnect()\
\
Disconnect from OPC Server");

PyDoc_STRVAR(OpcDA_additems_doc, "AddItem(item)\
\
Disconnect from OPC Server");


static OPCClient *_OPC = nullptr;

PyObject *OpcDA_setint(PyObject *self, PyObject *args, PyObject *kwargs) {
    /* Shared references that do not need Py_DECREF before returning. */
    int number = 0;

    /* Parse positional and keyword arguments */
    static char* keywords[] = { "obj", "number", NULL };
    if (!PyArg_ParseTuple(args, "i:setint", &number)) {
        return NULL;
    }
    __val = number;
    Py_RETURN_NONE;
}

PyObject* OpcDA_getint(PyObject* self) {
    return PyLong_FromLong(__val);
}

/*
 * List of functions to add to OpcDA in exec_OpcDA().
 */
static PyMethodDef OpcDA_functions[] = {
    { "setint", (PyCFunction)OpcDA_setint, METH_VARARGS, OpcDA_setint_doc },
    { "getint", (PyCFunction)OpcDA_getint, METH_NOARGS, OpcDA_getint_doc },
    { NULL, NULL, 0, NULL } /* marks end of array */
};

/*
 * Initialize OpcDA. May be called multiple times, so avoid
 * using static state.
 */
int exec_OpcDA(PyObject *module) {
    PyModule_AddFunctions(module, OpcDA_functions);

    PyModule_AddStringConstant(module, "__author__", "Roby Edrian");
    PyModule_AddStringConstant(module, "__version__", "1.0.0");
    PyModule_AddIntConstant(module, "year", 2022);

    return 0; /* success */
}

/*
 * Documentation for OpcDA.
 */
PyDoc_STRVAR(OpcDA_doc, "The OpcDA module");


static PyModuleDef_Slot OpcDA_slots[] = {
    { Py_mod_exec, exec_OpcDA },
    { 0, NULL }
};

static PyModuleDef OpcDA_def = {
    PyModuleDef_HEAD_INIT,
    "OpcDA",
    OpcDA_doc,
    0,              /* m_size */
    NULL,           /* m_methods */
    OpcDA_slots,
    NULL,           /* m_traverse */
    NULL,           /* m_clear */
    NULL,           /* m_free */
};

PyMODINIT_FUNC PyInit_OpcDA() {
    return PyModuleDef_Init(&OpcDA_def);
}
