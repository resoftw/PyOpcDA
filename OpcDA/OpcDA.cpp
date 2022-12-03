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

PyObject *OpcDA_init(PyObject *self, PyObject *args, PyObject *kwargs) {
    /* Shared references that do not need Py_DECREF before returning. */
    char* d = NULL, * u = NULL, * p = NULL;
    /* Parse positional and keyword arguments */
    if (!PyArg_ParseTuple(args, "sss:Init", &d,&u,&p)) {
        return PyLong_FromLong(-1);
    }
    return PyLong_FromLong(InitSecurity(d, u, p));
}

PyObject* OpcDA_connect(PyObject* self,PyObject *args) {
    char* h = NULL, * p = NULL;
    if (!PyArg_ParseTuple(args, "ss:Init", &h, &p)) {
        return PyLong_FromLong(-1);
    }
    wchar_t hst[256];
    wchar_t prg[256];
    mbstowcs(hst, h, 256);
    mbstowcs(prg, p, 256);
    if (_OPC == NULL) {
        _OPC = new OPCClient();
    }

    long r=_OPC->Connect(hst, prg);
    return PyLong_FromLong(r);
}

PyObject* OpcDA_disconnect(PyObject* self) {
    if (_OPC == NULL) return PyLong_FromLong(0);
    _OPC->Disconnect();
    Py_RETURN_NONE;
}

PyObject* OpcDA_additem(PyObject* self,PyObject *args) {

}
/*
 * List of functions to add to OpcDA in exec_OpcDA().
 */
static PyMethodDef OpcDA_functions[] = {
    { "Init", (PyCFunction)OpcDA_init, METH_VARARGS, OpcDA_init_doc },
    { "Connect", (PyCFunction)OpcDA_connect, METH_VARARGS, OpcDA_connect_doc },
    { "Disconnect", (PyCFunction)OpcDA_disconnect, METH_VARARGS, OpcDA_disconnect_doc },
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
