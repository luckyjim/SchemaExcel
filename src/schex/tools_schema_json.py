"""Parameters conversion and verification with a json schema"""

import json
import logging

from jsonschema import Draft4Validator, validators
from jsonschema.exceptions import ValidationError

s_logger = logging.getLogger(__name__)


def check_add_default_file(file_schema, d_pars):
    """
    Open json schema file and check dictionary d_pars, addition of default
    parameters for missing parameters

    :param file_schema: path and name of json schema file
    :type file_schema: str
    :param d_pars: dictionary of typed parameters
    :type d_pars: dict
    """
    with open(file_schema, "r") as fichier:
        json_schema = json.load(fichier)
    return check_add_default_dict(json_schema, d_pars)


def check_add_default_dict(schema, d_pars):
    """
    Check dictionary d_pars with schema json, addition of default
    parameters for missing parameters

    :param schema: json schema as dictionary
    :type schema: dict
    :param d_pars: dictionary of typed parameters
    :type d_pars: dict
    """
    for k in list(d_pars.keys()):
        if k not in list(schema["properties"].keys()):
            msg = f"\n\t#### Unknown parameter ['{k}'], not in json schema"
            msg += f" {sorted(list(schema['properties'].keys()))}"
            msg += f"\n\t#### Check the parameter name ['{k}'] in settings configuration file"
            s_logger.warning(msg)
            del d_pars[k]
    try:
        DefaultValidatingDraft4Validator(schema).validate(d_pars)
    except ValidationError as valid_err:
        s_logger.error(f"Validation KO: {valid_err}")
        return False
    s_logger.info("Parameters validated by json scheme")
    return True


def extend_with_default(validator_class):
    """
    https://python-jsonschema.readthedocs.io/en/stable/faq/ for more informations
    """
    validate_properties = validator_class.VALIDATORS["properties"]

    def set_defaults(validator, properties, instance, schema):
        """
        see https://python-jsonschema.readthedocs.io/en/stable/faq/ for more informations
        """
        #         for k in list(instance.keys()):
        #             if not k in list(properties.keys()):
        #                 del instance[k]
        #                 s_logger.warning(f"\n\t#### Key ['{k}'] not in schema keys {sorted(list(properties.keys()))}")
        for mproperty, subschema in properties.items():
            if "default" in subschema:
                instance.setdefault(mproperty, subschema["default"])

        for error in validate_properties(
            validator,
            properties,
            instance,
            schema,
        ):
            yield error

    return validators.extend(
        validator_class,
        {"properties": set_defaults},
    )


DefaultValidatingDraft4Validator = extend_with_default(Draft4Validator)
