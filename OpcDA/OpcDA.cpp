#include <Python.h>
#include <datetime.h>
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

PyDoc_STRVAR(OpcDA_additem_doc, "AddItem(item)\
\
Add Item");

PyDoc_STRVAR(OpcDA_readitem_doc, "ReadItem(item)\
\
Read Item \
item: item handle");

PyDoc_STRVAR(OpcDA_writeitem_doc, "WriteItem(item)\
\
Write Item \
item: item handle");

static OPCClient *_OPC = nullptr;
static int _nitems = 0;

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
    _nitems = 0;
    return PyLong_FromLong(r);
}

PyObject* OpcDA_disconnect(PyObject* self) {
    if (_OPC == NULL) return PyLong_FromLong(0);
    _OPC->Disconnect();
    Py_RETURN_NONE;
}

PyObject* OpcDA_additem(PyObject* self,PyObject *args) {
    if (!_OPC) return PyTuple_Pack(2, PyBool_FromLong(0), Py_None);
    char* n = NULL;
    if (!PyArg_ParseTuple(args, "s:AddItem", &n)) {
        return PyTuple_Pack(2, PyBool_FromLong(0), PyUnicode_FromString("ParseTUpler!"));
    }
    OpcItem o;
    if (_OPC->AddItem(_nitems+1, std::string(n), VT_EMPTY, o)==S_OK) {
        _nitems++;
        return PyTuple_Pack(2, PyBool_FromLong(1),
            PyTuple_Pack(3,
                PyUnicode_FromString(VarTypeStr(o.DataType).c_str()),
                PyUnicode_FromString(o.ItemID.c_str()), 
                PyLong_FromLong(o.Handle)));
    }
    else return PyTuple_Pack(2, PyBool_FromLong(0), PyUnicode_FromString("Error adding!"));
}

PyObject* OpcDA_readitem(PyObject* self, PyObject* args) {
    OPCHANDLE opch = 0;
    if (!PyArg_ParseTuple(args, "k:ReadItem", &opch))
    {
        return PyTuple_Pack(2, PyBool_FromLong(0), PyUnicode_FromString("Invalid parameter!"));
    }
    OpcData d;
    HRESULT h = _OPC->ReadItem(opch, d);
    if (h == S_OK) {
        return PyTuple_Pack(2, PyBool_FromLong(1), PyFloat_FromDouble(d.Value.dblVal));
    }
    else
        return PyTuple_Pack(2, PyBool_FromLong(0), PyLong_FromLong(h));
}

PyObject* OpcDA_writeitem(PyObject* self, PyObject* args) {
    OPCHANDLE opch = 0;
    double v = 0.0;
    if (!PyArg_ParseTuple(args, "kd:ReadItem", &opch,&v))
    {
        return PyTuple_Pack(2, PyBool_FromLong(0), PyUnicode_FromString("Invalid parameter!"));
    }
    VARIANT val{};
    val.dblVal=v;
    val.vt = VT_R8;
    //printf("Value: %g\n", v);
    HRESULT h = _OPC->Write(opch, val);
    if (h == S_OK) {
        return PyTuple_Pack(2, PyBool_FromLong(1), Py_None);
    }
    else
        return PyTuple_Pack(2, PyBool_FromLong(0), PyLong_FromLong(h));
}

/*
 * List of functions to add to OpcDA in exec_OpcDA().
 */
static PyMethodDef OpcDA_functions[] = {
    { "Init", (PyCFunction)OpcDA_init, METH_VARARGS, OpcDA_init_doc },
    { "Connect", (PyCFunction)OpcDA_connect, METH_VARARGS, OpcDA_connect_doc },
    { "Disconnect", (PyCFunction)OpcDA_disconnect, METH_VARARGS, OpcDA_disconnect_doc },
    { "AddItem", (PyCFunction)OpcDA_additem, METH_VARARGS, OpcDA_additem_doc },
    { "ReadItem", (PyCFunction)OpcDA_readitem, METH_VARARGS, OpcDA_readitem_doc },
    { "WriteItem", (PyCFunction)OpcDA_writeitem, METH_VARARGS, OpcDA_writeitem_doc },
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
    PyDateTime_IMPORT;
    return PyModuleDef_Init(&OpcDA_def);
}
