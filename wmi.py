# -*- coding: utf-8 -*-
from __future__ import with_statement
__version__ = "1.5.1"
__VERSION__ = __version__
__title__ = "WMI"
__description__ = "Windows Management Instrumentation"
__author__ = "Tim Golden"
__email__ = "mail@timgolden.me.uk"
__url__ = "http://timgolden.me.uk/python/wmi.html"
__license__ = "http://www.opensource.org/licenses/mit-license.php"
_DEBUG = False

import sys
import csv
import datetime
import re
import struct
import warnings

from win32com.client import GetObject, Dispatch
import pywintypes

def signed_to_unsigned(signed):
    unsigned, = struct.unpack("L", struct.pack("l", signed))
    return unsigned

class SelfDeprecatingDict(object):

    dict_only = set(dir(dict)).difference(dir(list))

    def __init__(self, dictlike):
        self.dict = dict(dictlike)
        self.list = list(self.dict)

    def __getattr__(self, attribute):
        if attribute in self.dict_only:
            warnings.warn("In future this will be a list and not a dictionary", DeprecationWarning)
            return getattr(self.dict, attribute)
        else:
            return getattr(self.list, attribute)

    def __iter__(self):
        return iter(self.list)

    def __str__(self):
        return str(self.list)

    def __repr__(self):
        return repr(self.list)

    def __getitem__(self, item):
        try:
            return self.list[item]
        except TypeError:
            return self.dict[item]

class ProvideConstants(object):
    def __init__(self, comobj):
        comobj.__dict__["_constants"] = self
        self.__typecomp = \
        comobj._oleobj_.GetTypeInfo().GetContainingTypeLib()[0].GetTypeComp()

    def __getattr__(self, name):
        if name.startswith("__") and name.endswith("__"):
         raise AttributeError(name)
        result = self.__typecomp.Bind(name)
        if not result[0]:
         raise AttributeError(name)
        return result[1].value

obj = GetObject("winmgmts:")
ProvideConstants(obj)

wbemErrInvalidQuery = obj._constants.wbemErrInvalidQuery
wbemErrTimedout = obj._constants.wbemErrTimedout
wbemFlagReturnImmediately = obj._constants.wbemFlagReturnImmediately
wbemFlagForwardOnly = obj._constants.wbemFlagForwardOnly

#
# Exceptions
#
class x_wmi(Exception):
    def __init__(self, info="", com_error=None):
        self.info = info
        self.com_error = com_error

    def __str__(self):
        return "<x_wmi: %s %s>" %(
            self.info or "Unexpected COM Error",
            self.com_error or "(no underlying exception)"
        )

class x_wmi_invalid_query(x_wmi):
    pass

class x_wmi_timed_out(x_wmi):
    pass

class x_wmi_no_namespace(x_wmi):
    pass

class x_access_denied(x_wmi):
    pass

class x_wmi_authentication(x_wmi):
    pass

class x_wmi_uninitialised_thread(x_wmi):
    pass

WMI_EXCEPTIONS = {
    signed_to_unsigned(wbemErrInvalidQuery) : x_wmi_invalid_query,
    signed_to_unsigned(wbemErrTimedout) : x_wmi_timed_out,
    0x80070005 : x_access_denied,
    0x80041003 : x_access_denied,
    0x800401E4 : x_wmi_uninitialised_thread,
}

def handle_com_error(err=None):
    if err is None:
        _, err, _ = sys.exc_info()
    hresult_code, hresult_name, additional_info, parameter_in_error = err.args
    hresult_code = signed_to_unsigned(hresult_code)
    exception_string = ["%s - %s" %(hex(hresult_code), hresult_name)]
    scode = None
    if additional_info:
        wcode, source_of_error, error_description, whlp_file, whlp_context, scode = additional_info
        scode = signed_to_unsigned(scode)
        exception_string.append("    Error in: %s" % source_of_error)
        exception_string.append("    %s - %s" %(hex(scode),(error_description or "").strip()))
    for error_code, klass in WMI_EXCEPTIONS.items():
        if error_code in(hresult_code, scode):
            break
    else:
        klass = x_wmi
    raise klass(com_error=err)


BASE = datetime.datetime(1601, 1, 1)
def from_1601(ns100):
    return BASE + datetime.timedelta(microseconds=int(ns100) / 10)

def from_time(year=None, month=None, day=None, hours=None, minutes=None, seconds=None, microseconds=None, timezone=None):
    def str_or_stars(i, length):
        if i is None:
            return "*" * length
        else:
            return str(i).rjust(length, "0")

    wmi_time = ""
    wmi_time += str_or_stars(year, 4)
    wmi_time += str_or_stars(month, 2)
    wmi_time += str_or_stars(day, 2)
    wmi_time += str_or_stars(hours, 2)
    wmi_time += str_or_stars(minutes, 2)
    wmi_time += str_or_stars(seconds, 2)
    wmi_time += "."
    wmi_time += str_or_stars(microseconds, 6)
    if timezone is None:
        wmi_time += "+"
    else:
        try:
            int(timezone)
        except ValueError:
            wmi_time += "+"
        else:
            if timezone >= 0:
                wmi_time += "+"
            else:
                wmi_time += "-"
                timezone = abs(timezone)
    wmi_time += str_or_stars(timezone, 3)

    return wmi_time

def to_time(wmi_time):
    def int_or_none(s, start, end):
        try:
            return int(s[start:end])
        except ValueError:
            return None

    year = int_or_none(wmi_time, 0, 4)
    month = int_or_none(wmi_time, 4, 6)
    day = int_or_none(wmi_time, 6, 8)
    hours = int_or_none(wmi_time, 8, 10)
    minutes = int_or_none(wmi_time, 10, 12)
    seconds = int_or_none(wmi_time, 12, 14)
    microseconds = int_or_none(wmi_time, 15, 21)
    timezone = wmi_time[22:]
    if timezone == "***":
        timezone = None

    return year, month, day, hours, minutes, seconds, microseconds, timezone

def _set(obj, attribute, value):
    obj.__dict__[attribute] = value

class _wmi_method(object):
    def __init__(self, ole_object, method_name):
        """
        :param ole_object: The WMI class/instance whose method is to be called
        :param method_name: The name of the method to be called
        """

        #
        # FIXME: make use of this function, copied from a defunct branch
        #
        def parameter_names(method_parameters):
            parameter_names = []
            for param in method_parameters.Properties_:
                name, is_array = param.Name, param.IsArray
                datatype = bitmap = None
                for qualifier in param.Qualifiers_:
                    if qualifier.Name == "CIMTYPE":
                        datatype = qualifier.Value
                    elif qualifier.Name == "BitMap":
                        bitmap = [int(b) for b in qualifier.Value]
                parameter_names.append((name, is_array, datatype, bitmap))
            return parameter_names

        try:
            self.ole_object = Dispatch(ole_object)
            self.method = ole_object.Methods_(method_name)
            self.qualifiers = {}
            for q in self.method.Qualifiers_:
                self.qualifiers[q.Name] = q.Value
            self.provenance = "\n".join(self.qualifiers.get("MappingStrings", []))

            self.in_parameters = self.method.InParameters
            self.out_parameters = self.method.OutParameters
            if self.in_parameters is None:
                self.in_parameter_names = []
            else:
                self.in_parameter_names = [(i.Name, i.IsArray) for i in self.in_parameters.Properties_]
            if self.out_parameters is None:
                self.out_parameter_names = []
            else:
                self.out_parameter_names = [(i.Name, i.IsArray) for i in self.out_parameters.Properties_]

            doc = "%s (%s) => (%s)" % (
                method_name,
                ", ".join([name +("", "[]")[is_array] for (name, is_array) in self.in_parameter_names]),
                ", ".join([name +("", "[]")[is_array] for (name, is_array) in self.out_parameter_names])
            )
            privileges = self.qualifiers.get("Privileges", [])
            if privileges:
                doc += " | Needs: " + ", ".join(privileges)
            self.__doc__ = doc
        except pywintypes.com_error:
            handle_com_error()

    def __call__(self, *args, **kwargs):
        try:
            if self.in_parameters:
                parameter_names = {}
                for name, is_array in self.in_parameter_names:
                    parameter_names[name] = is_array

                parameters = self.in_parameters

                #
                # Check positional parameters first
                #
                for n_arg in range(len(args)):
                    arg = args[n_arg]
                    parameter = parameters.Properties_[n_arg]
                    if parameter.IsArray:
                        try: list(arg)
                        except TypeError: raise TypeError("parameter %d must be iterable" % n_arg)
                    parameter.Value = arg

                #
                # If any keyword param supersedes a positional one,
                # it'll simply overwrite it.
                #
                for k, v in kwargs.items():
                    is_array = parameter_names.get(k)
                    if is_array is None:
                        raise AttributeError("%s is not a valid parameter for %s" %(k, self.__doc__))
                    else:
                        if is_array:
                            try: list(v)
                            except TypeError: raise TypeError("%s must be iterable" % k)
                    parameters.Properties_(k).Value = v

                result = self.ole_object.ExecMethod_(self.method.Name, self.in_parameters)
            else:
                result = self.ole_object.ExecMethod_(self.method.Name)

            results = []
            for name, is_array in self.out_parameter_names:
                value = result.Properties_(name).Value
                if is_array:
                    #
                    # Thanks to Jonas Bjering for bug report and patch
                    #
                    results.append(list(value or []))
                else:
                    results.append(value)
            return tuple(results)

        except pywintypes.com_error:
            handle_com_error()

    def __repr__(self):
        return "<function %s>" % self.__doc__

class _wmi_property(object):

    def __init__(self, property):
        self.property = property
        self.name = property.Name
        self.value = property.Value
        self.qualifiers = dict((q.Name, q.Value) for q in property.Qualifiers_)
        self.type = self.qualifiers.get("CIMTYPE", None)
        self.provenance = "\n".join(self.qualifiers.get("MappingStrings", []))

    def set(self, value):
        self.property.Value = value

    def __repr__(self):
        return "<wmi_property: %s>" % self.name

    def __getattr__(self, attr):
        return getattr(self.property, attr)

#
# class _wmi_object
#
class _wmi_object(object):
    def __init__(self, ole_object, instance_of=None, fields=[], property_map={}):
        try:
            _set(self, "ole_object", ole_object)
            _set(self, "id", ole_object.Path_.DisplayName.lower())
            _set(self, "_instance_of", instance_of)
            _set(self, "properties", {})
            _set(self, "methods", {})
            _set(self, "property_map", property_map)
            _set(self, "_associated_classes", None)
            _set(self, "_keys", None)

            if fields:
                for field in fields:
                    self.properties[field] = None
            else:
                for p in ole_object.Properties_:
                    self.properties[p.Name] = None

            for m in ole_object.Methods_:
                self.methods[m.Name] = None

            _set(self, "_properties", self.properties.keys())
            _set(self, "_methods", self.methods.keys())
            _set(self, "qualifiers", dict((q.Name, q.Value) for q in self.ole_object.Qualifiers_))
            _set(self, "is_association", "Association" in self.qualifiers)

        except pywintypes.com_error:
            handle_com_error()

    def __lt__(self, other):
        return self.id < other.id

    def __str__(self):
        try:
            return self.ole_object.GetObjectText_()
        except pywintypes.com_error:
            handle_com_error()

    def __repr__(self):
        try:
            return "<%s: %s>" % (self.__class__.__name__, self.Path_.Path.encode("ascii", "backslashreplace"))
        except pywintypes.com_error:
            handle_com_error()

    def _cached_properties(self, attribute):
        if self.properties[attribute] is None:
            self.properties[attribute] = _wmi_property(self.ole_object.Properties_(attribute))
        return self.properties[attribute]

    def _cached_methods(self, attribute):
        if self.methods[attribute] is None:
            self.methods[attribute] = _wmi_method(self.ole_object, attribute)
        return self.methods[attribute]

    def __getattr__(self, attribute):
        try:
            if attribute in self.properties:
                property = self._cached_properties(attribute)
                factory = self.property_map.get(attribute, self.property_map.get(property.type, lambda x: x))
                value = factory(property.value)
                #
                # If this is an association, certain of its properties
                # are actually the paths to the aspects of the association,
                # so translate them automatically into WMI objects.
                #
                if property.type.startswith("ref:"):
                    return WMI(moniker=value)
                else:
                    return value
            elif attribute in self.methods:
                return self._cached_methods(attribute)
            else:
                return getattr(self.ole_object, attribute)
        except pywintypes.com_error:
            handle_com_error()

    def __setattr__(self, attribute, value):
        try:
            if attribute in self.properties:
                self._cached_properties(attribute).set(value)
                if self.ole_object.Path_.Path:
                    self.ole_object.Put_()
            else:
                raise AttributeError(attribute)
        except pywintypes.com_error:
            handle_com_error()

    def __eq__(self, other):
        try:
            return self.id == other.id
        except AttributeError:
            return False

    def __hash__(self):
        return hash(self.id)

    def _getAttributeNames(self):
         attribs = [str(x) for x in self.methods.keys()]
         attribs.extend([str(x) for x in self.properties.keys()])
         return attribs

    def _get_keys(self):
        # NB You can get the keys of an instance more directly, via
        # Path\_.Keys but this doesn't apply to classes. The technique
        # here appears to work for both.
        if self._keys is None:
            _set(self, "_keys", [])
            for property in self.ole_object.Properties_:
                for qualifier in property.Qualifiers_:
                    if qualifier.Name == "key" and qualifier.Value:
                        self._keys.append(property.Name)
        return self._keys
    keys = property(_get_keys)

    def wmi_property(self, property_name):
        return _wmi_property(self.ole_object.Properties_(property_name))

    def put(self):
        self.ole_object.Put_()

    def set(self, **kwargs):
        if kwargs:
            try:
                for attribute, value in kwargs.items():
                    if attribute in self.properties:
                        self._cached_properties(attribute).set(value)
                    else:
                        raise AttributeError(attribute)
                #
                # Only try to write the attributes
                #    back if the object exists.
                #
                if self.ole_object.Path_.Path:
                    self.ole_object.Put_()
            except pywintypes.com_error:
                handle_com_error()

    def path(self):
        try:
            return self.ole_object.Path_
        except pywintypes.com_error:
            handle_com_error()

    def derivation(self):
        try:
            return self.ole_object.Derivation_
        except pywintypes.com_error:
            handle_com_error()

    def _cached_associated_classes(self):
        if isinstance(self, _wmi_class):
            obj = self
        else:
            obj = self._instance_of
        if obj._associated_classes is None:
            try:
                associated_classes = dict(
                   (assoc.Path_.Class, _wmi_class(self._namespace, assoc)) for
                        assoc in obj.ole_object.Associators_(bSchemaOnly=True)
                )
                _set(obj, "_associated_classes", associated_classes)
            except pywintypes.com_error:
                handle_com_error()

        return obj._associated_classes
    associated_classes = property(_cached_associated_classes)

    def associators(self, wmi_association_class="", wmi_result_class=""):
        try:
            return [
                _wmi_object(i) for i in \
                    self.ole_object.Associators_(
                     strAssocClass=wmi_association_class,
                     strResultClass=wmi_result_class
                 )
            ]
        except pywintypes.com_error:
            handle_com_error()

    def references(self, wmi_class=""):
        try:
            return [_wmi_object(i) for i in self.ole_object.References_(strResultClass=wmi_class)]
        except pywintypes.com_error:
            handle_com_error()

class _wmi_event(_wmi_object):
    event_type_re = re.compile("__Instance(Creation|Modification|Deletion)Event")
    def __init__(self, event, event_info, fields=[]):
        _wmi_object.__init__(self, event, fields=fields)
        _set(self, "event_type", None)
        _set(self, "timestamp", None)
        _set(self, "previous", None)

        if event_info:
            event_type = self.event_type_re.match(event_info.Path_.Class).group(1).lower()
            _set(self, "event_type", event_type)
            if hasattr(event_info, "TIME_CREATED"):
                _set(self, "timestamp", from_1601(event_info.TIME_CREATED))
            if hasattr(event_info, "PreviousInstance"):
                _set(self, "previous", event_info.PreviousInstance)

#
# class _wmi_class
#
class _wmi_class(_wmi_object):
    def __init__(self, namespace, wmi_class):
        _wmi_object.__init__(self, wmi_class)
        _set(self, "_class_name", wmi_class.Path_.Class)
        if namespace:
            _set(self, "_namespace", namespace)
        else:
            class_moniker = wmi_class.Path_.DisplayName
            winmgmts, namespace_moniker, class_name = class_moniker.split(":")
            namespace = _wmi_namespace(GetObject(winmgmts + ":" + namespace_moniker), False)
            _set(self, "_namespace", namespace)

    def __getattr__(self, attribute):
        try:
            if attribute in self.properties:
                return _wmi_property(self.Properties_(attribute))
            else:
                return _wmi_object.__getattr__(self, attribute)
        except pywintypes.com_error:
            handle_com_error()


    def to_csv(self, filepath=None):
        def _to_utf8(item):
            if isinstance(item, unicode):
                return item.encode("utf-8")
            else:
                return str(item)
        if filepath is None:
            filepath = self._class_name + ".csv"
        fields = list(p.Name for p in self.ole_object.Properties_)
        with open(filepath, "wb") as f:
            writer = csv.writer(f)
            writer.writerow(fields)
            for instance in self.query():
                writer.writerow([_to_utf8(getattr(instance, field)) for field in fields])

    def query(self, fields=[], **where_clause):
        if self._namespace is None:
            raise x_wmi_no_namespace("You cannot query directly from a WMI class")

        try:
            field_list = ", ".join(fields) or "*"
            wql = "SELECT " + field_list + " FROM " + self._class_name
            if where_clause:
                wql += " WHERE " + " AND ". join(["%s = %r" % (k, str(v)) for k, v in where_clause.items()])
            return self._namespace.query(wql, self, fields)
        except pywintypes.com_error:
            handle_com_error()

    __call__ = query

    def watch_for(
        self,
        notification_type="operation",
        delay_secs=1,
        fields=[],
        **where_clause
    ):
        if self._namespace is None:
            raise x_wmi_no_namespace("You cannot watch directly from a WMI class")

        valid_notification_types = ("operation", "creation", "deletion", "modification")
        if notification_type.lower () not in valid_notification_types:
            raise x_wmi ("notification_type must be one of %s" % ", ".join (valid_notification_types))

        return self._namespace.watch_for(
            notification_type=notification_type,
            wmi_class=self,
            delay_secs=delay_secs,
            fields=fields,
            **where_clause
        )

    def instances(self):
        try:
            return [_wmi_object(instance, self) for instance in self.Instances_()]
        except pywintypes.com_error:
            handle_com_error()

    def new(self, **kwargs):
        try:
            obj = _wmi_object(self.SpawnInstance_(), self)
            obj.set(**kwargs)
            return obj
        except pywintypes.com_error:
            handle_com_error()

#
# class _wmi_result
#
class _wmi_result(object):
    def __init__(self, obj, attributes):
        if attributes:
            for attr in attributes:
                self.__dict__[attr] = obj.Properties_(attr).Value
        else:
            for p in obj.Properties_:
                attr = p.Name
                self.__dict__[attr] = obj.Properties_(attr).Value

#
# class WMI
#
class _wmi_namespace(object):
    def __init__(self, namespace, find_classes):
        _set(self, "_namespace", namespace)
        #
        # wmi attribute preserved for backwards compatibility
        #
        _set(self, "wmi", namespace)

        self._classes = None
        self._classes_map = {}
        if find_classes:
            _ = self.classes

    def __repr__(self):
        return "<_wmi_namespace: %s>" % self.wmi

    def __str__(self):
        return repr(self)

    def _get_classes(self):
        if self._classes is None:
            self._classes = self.subclasses_of()
        return SelfDeprecatingDict(dict.fromkeys(self._classes))
    classes = property(_get_classes)

    def get(self, moniker):
        try:
            return _wmi_object(self.wmi.Get(moniker))
        except pywintypes.com_error:
            handle_com_error()

    def handle(self):
        return self._namespace

    def subclasses_of(self, root="", regex=r".*"):
        try:
            SubclassesOf = self._namespace.SubclassesOf
        except AttributeError:
            return set()
        else:
            return set(
                c.Path_.Class
                    for c in SubclassesOf(root)
                    if re.match(regex, c.Path_.Class)
            )

    def instances(self, class_name):
        try:
            return [_wmi_object(obj) for obj in self._namespace.InstancesOf(class_name)]
        except pywintypes.com_error:
            handle_com_error()

    def new(self, wmi_class, **kwargs):
        """This is now implemented by a call to :meth:`_wmi_class.new`"""
        return getattr(self, wmi_class).new(**kwargs)

    new_instance_of = new

    def _raw_query(self, wql):
        flags = wbemFlagReturnImmediately | wbemFlagForwardOnly
        wql = wql.replace("\\", "\\\\")
        try:
            return self._namespace.ExecQuery(strQuery=wql, iFlags=flags)
        except pywintypes.com_error:
            handle_com_error()

    def query(self, wql, instance_of=None, fields=[]):
        return [ _wmi_object(obj, instance_of, fields) for obj in self._raw_query(wql) ]

    def fetch_as_classes(self, wmi_classname, fields=(), **where_clause):
        wql = "SELECT %s FROM %s" %(fields and ", ".join(fields) or "*", wmi_classname)
        if where_clause:
            wql += " WHERE " + " AND ".join(["%s = '%s'" %(k, v) for k, v in where_clause.items()])
        return [_wmi_result(obj, fields) for obj in self._raw_query(wql)]

    def fetch_as_lists(self, wmi_classname, fields, **where_clause):
        wql = "SELECT %s FROM %s" %(", ".join(fields), wmi_classname)
        if where_clause:
            wql += " WHERE " + " AND ".join(["%s = '%s'" %(k, v) for k, v in where_clause.items()])
        results = []
        for obj in self._raw_query(wql):
                results.append([obj.Properties_(field).Value for field in fields])
        return results

    def watch_for(
        self,
        raw_wql=None,
        notification_type="operation",
        wmi_class=None,
        delay_secs=1,
        fields=[],
        **where_clause
    ):
        if raw_wql:
            wql = raw_wql
            is_extrinsic = False
        else:
            if isinstance(wmi_class, _wmi_class):
                class_name = wmi_class._class_name
            else:
                class_name = wmi_class
                wmi_class = getattr(self, class_name)
            is_extrinsic = "__ExtrinsicEvent" in wmi_class.derivation()
            fields = set(['TargetInstance'] + (fields or ["*"]))
            field_list = ", ".join(fields)
            if is_extrinsic:
                if where_clause:
                    where = " WHERE " + " AND ".join(["%s = '%s'" %(k, v) for k, v in where_clause.items()])
                else:
                    where = ""
                wql = "SELECT " + field_list + " FROM " + class_name + where
            else:
                if where_clause:
                    where = " AND " + " AND ".join(["TargetInstance.%s = '%s'" %(k, v) for k, v in where_clause.items()])
                else:
                    where = ""
                wql = \
                    "SELECT %s FROM __Instance%sEvent WITHIN %d WHERE TargetInstance ISA '%s' %s" % \
                   (field_list, notification_type, delay_secs, class_name, where)

        try:
            return _wmi_watcher(
                self._namespace.ExecNotificationQuery(wql),
                is_extrinsic=is_extrinsic,
                fields=fields
            )
        except pywintypes.com_error:
            handle_com_error()

    def __getattr__(self, attribute):
        try:
            return self._cached_classes(attribute)
        except pywintypes.com_error:
            return getattr(self._namespace, attribute)

    def _cached_classes(self, class_name):
        if class_name not in self._classes_map:
            self._classes_map[class_name] = _wmi_class(self, self._namespace.Get(class_name))
        return self._classes_map[class_name]

    def _getAttributeNames(self):
        return [x for x in self.classes if not x.startswith('__')]

#
# class _wmi_watcher
#
class _wmi_watcher(object):
    _event_property_map = {
        "TargetInstance" : _wmi_object,
        "PreviousInstance" : _wmi_object
    }
    def __init__(self, wmi_event, is_extrinsic, fields=[]):
        self.wmi_event = wmi_event
        self.is_extrinsic = is_extrinsic
        self.fields = fields

    def __call__(self, timeout_ms=-1):
        try:
            event = self.wmi_event.NextEvent(timeout_ms)
            if self.is_extrinsic:
                return _wmi_event(event, None, self.fields)
            else:
                return _wmi_event(
                    event.Properties_("TargetInstance").Value,
                    _wmi_object(event, property_map=self._event_property_map),
                    self.fields
                )
        except pywintypes.com_error:
            handle_com_error()

PROTOCOL = "winmgmts:"
def connect(
    computer="",
    impersonation_level="",
    authentication_level="",
    authority="",
    privileges="",
    moniker="",
    wmi=None,
    namespace="",
    suffix="",
    user="",
    password="",
    find_classes=False,
    debug=False
):
    global _DEBUG
    _DEBUG = debug

    try:
        try:
            if wmi:
                obj = wmi

            elif moniker:
                if not moniker.startswith(PROTOCOL):
                    moniker = PROTOCOL + moniker
                obj = GetObject(moniker)

            else:
                if user:
                    if privileges or suffix:
                        raise x_wmi_authentication("You can't specify privileges or a suffix as well as a username")
                    elif computer in(None, '', '.'):
                        raise x_wmi_authentication("You can only specify user/password for a remote connection")
                    else:
                        obj = connect_server(
                            server=computer,
                            namespace=namespace,
                            user=user,
                            password=password,
                            authority=authority,
                            impersonation_level=impersonation_level,
                            authentication_level=authentication_level
                        )

                else:
                    moniker = construct_moniker(
                        computer=computer,
                        impersonation_level=impersonation_level,
                        authentication_level=authentication_level,
                        authority=authority,
                        privileges=privileges,
                        namespace=namespace,
                        suffix=suffix
                    )
                    obj = GetObject(moniker)

            wmi_type = get_wmi_type(obj)

            if wmi_type == "namespace":
                return _wmi_namespace(obj, find_classes)
            elif wmi_type == "class":
                return _wmi_class(None, obj)
            elif wmi_type == "instance":
                return _wmi_object(obj)
            else:
                raise x_wmi("Unknown moniker type")

        except pywintypes.com_error:
            handle_com_error()

    except x_wmi_uninitialised_thread:
        raise x_wmi_uninitialised_thread("WMI returned a syntax error: you're probably running inside a thread without first calling pythoncom.CoInitialize[Ex]")

WMI = connect

def construct_moniker(
    computer=None,
    impersonation_level=None,
    authentication_level=None,
    authority=None,
    privileges=None,
    namespace=None,
    suffix=None
):
    security = []
    if impersonation_level: security.append("impersonationLevel=%s" % impersonation_level)
    if authentication_level: security.append("authenticationLevel=%s" % authentication_level)
    #
    # Use of the authority descriptor is invalid on the local machine
    #
    if authority and computer: security.append("authority=%s" % authority)
    if privileges: security.append("(%s)" % ", ".join(privileges))

    moniker = [PROTOCOL]
    if security: moniker.append("{%s}!" % ",".join(security))
    if computer: moniker.append("//%s/" % computer)
    if namespace:
        parts = re.split(r"[/\\]", namespace)
        if parts[0] != 'root':
            parts.insert(0, "root")
        moniker.append("/".join(parts))
    if suffix: moniker.append(":%s" % suffix)
    return "".join(moniker)

def get_wmi_type(obj):
    try:
        path = obj.Path_
    except AttributeError:
        return "namespace"
    else:
        if path.IsClass:
            return "class"
        else:
            return "instance"

def connect_server(
    server,
    namespace = "",
    user = "",
    password = "",
    locale = "",
    authority = "",
    impersonation_level="",
    authentication_level="",
    security_flags = 0x80,
    named_value_set = None
):
    if impersonation_level:
        try:
            impersonation = getattr(obj._constants, "wbemImpersonationLevel%s" % impersonation_level.title())
        except AttributeError:
            raise x_wmi_authentication("No such impersonation level: %s" % impersonation_level)
    else:
        impersonation = None

    if authentication_level:
        try:
            authentication = getattr(obj._constants, "wbemAuthenticationLevel%s" % authentication_level.title())
        except AttributeError:
            raise x_wmi_authentication("No such impersonation level: %s" % impersonation_level)
    else:
        authentication = None

    server = Dispatch("WbemScripting.SWbemLocator").\
        ConnectServer(
            server,
            namespace,
            user,
            password,
            locale,
            authority,
            security_flags,
            named_value_set
        )
    if impersonation:
        server.Security_.ImpersonationLevel    = impersonation
    if authentication:
        server.Security_.AuthenticationLevel    = authentication
    return server

def Registry(
    computer=None,
    impersonation_level="Impersonate",
    authentication_level="Default",
    authority=None,
    privileges=None,
    moniker=None
):

    warnings.warn("This function can be implemented using wmi.WMI(namespace='DEFAULT').StdRegProv", DeprecationWarning)
    if not moniker:
        moniker = construct_moniker(
            computer=computer,
            impersonation_level=impersonation_level,
            authentication_level=authentication_level,
            authority=authority,
            privileges=privileges,
            namespace="default",
            suffix="StdRegProv"
        )

    try:
        return _wmi_object(GetObject(moniker))

    except pywintypes.com_error:
        handle_com_error()

if __name__ == '__main__':
    system = WMI()
    for my_computer in system.Win32_ComputerSystem():
        print("Disks on %s" % my_computer.Name)
        for disk in system.Win32_LogicalDisk():
            print("%s; %s; %s" % (disk.Caption, disk.Description, disk.ProviderName or ""))