#!/usr/bin/env python3
"""
Обратная совместимость: реэкспорт auth-модулей.
Новый код использует cischecker.auth.* напрямую.
"""
from cischecker.auth.jwt_flow import auth_jwt, set_log_fn
from cischecker.auth.certificates import list_certificates
from cischecker.auth.signer import sign_data

__all__ = ["auth_jwt", "list_certificates", "sign_data", "set_log_fn"]