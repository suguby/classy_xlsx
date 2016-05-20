# -*- coding: utf-8 -*-
from collections import OrderedDict
import copy

from bunch import Bunch


class XlsxContext(object):
    default_context = None
    PARENT_ATTR_NAME = 'nope'

    def __init__(self, context=None):
        if self.default_context:
            self._context = self.default_context.copy()
        else:
            self._context = Bunch()
        if context:
            self._context.update(context)

    def get_extra_context(self):
        return dict()

    @property
    def parent(self):
        try:
            return getattr(self, self.PARENT_ATTR_NAME)
        except AttributeError:
            return None

    @property
    def context(self):
        context = Bunch(self.parent.context.copy()) if self.parent else Bunch()
        if self._context:
            context.update(self._context.copy())
        extra_context = self.get_extra_context()
        if extra_context:
            context.update(extra_context.copy())
        return context

    def update_context(self, **kwargs):
        self._context.update(kwargs)


class XlsxField(XlsxContext):
    __position = 0

    def __init__(self, **kwargs):
        super(XlsxField, self).__init__(**kwargs)
        self.name = None
        XlsxField.__position += 1
        self.position = XlsxField.__position

    @classmethod
    def copy_fields_to_instance(cls, instance, **attrs):
        fields = []
        for name in dir(instance.__class__):
            if name.startswith('_'):
                continue
            class_field = getattr(instance.__class__, name)
            if isinstance(class_field, cls):
                instance_field = copy.deepcopy(class_field)
                instance_field.name = name
                setattr(instance_field, cls.PARENT_ATTR_NAME, instance)
                for k, v in attrs.iteritems():
                    setattr(instance_field, k, v)
                setattr(instance, name, instance_field)
                fields.append(instance_field)
            if isinstance(class_field, tuple) and len(class_field) == 1 and isinstance(class_field[0], cls):
                print "Probably {} in {} has unnecessary comma!".format(name, instance.__name__)
        fields.sort(key=lambda x: x.position)
        for i, field in enumerate(fields):
            field.position = i
        return OrderedDict((field.name, field) for field in fields)


class ClassyXlsxException(Exception):
    pass

