"""Helper for reading Samsara Function secrets at runtime.

Mirrors the pattern from Samsara's official functions-examples repo
(basic/just-secrets): assume the Function's execution role via STS, then
read the JSON secrets blob from SSM Parameter Store. boto3 is available
in the Samsara Functions runtime; no extra dependencies are needed.
"""

import json
import os
from datetime import datetime, timedelta, timezone

import boto3

_CREDENTIALS = None
_CREDENTIALS_EXPIRY = None
_SECRETS = None

_REFRESH_MARGIN = timedelta(minutes=20)


def get_credentials(force_refresh=False):
    global _CREDENTIALS, _CREDENTIALS_EXPIRY
    now = datetime.now(timezone.utc)
    if (
        not force_refresh
        and _CREDENTIALS is not None
        and _CREDENTIALS_EXPIRY is not None
        and now + _REFRESH_MARGIN < _CREDENTIALS_EXPIRY
    ):
        return _CREDENTIALS

    sts = boto3.client("sts")
    res = sts.assume_role(
        RoleArn=os.environ["SamsaraFunctionExecRoleArn"],
        RoleSessionName=os.environ["SamsaraFunctionName"],
    )
    creds = res["Credentials"]
    _CREDENTIALS = {
        "aws_access_key_id": creds["AccessKeyId"],
        "aws_secret_access_key": creds["SecretAccessKey"],
        "aws_session_token": creds["SessionToken"],
    }
    expiry = creds["Expiration"]
    if expiry.tzinfo is None:
        expiry = expiry.replace(tzinfo=timezone.utc)
    _CREDENTIALS_EXPIRY = expiry
    return _CREDENTIALS


def get_secrets(force_refresh=False):
    """Return the Function's configured secrets as a dict."""
    global _SECRETS
    if not force_refresh and _SECRETS is not None:
        return _SECRETS

    ssm = boto3.client("ssm", **get_credentials(force_refresh))
    value = ssm.get_parameter(
        Name=os.environ["SamsaraFunctionSecretsPath"],
        WithDecryption=True,
    )["Parameter"]["Value"]
    _SECRETS = json.loads(value)
    return _SECRETS


def apply_to_env(secrets):
    """Copy secrets into os.environ for libraries that read env vars."""
    for key, value in secrets.items():
        os.environ[key] = str(value)
