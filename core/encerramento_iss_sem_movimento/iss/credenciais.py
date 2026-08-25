from pydantic import BaseModel, SecretStr, field_validator
import re


class CredenciaisISSFortaleza(BaseModel):
    cpf: SecretStr
    senha: SecretStr

    @field_validator("cpf")
    @classmethod
    def validar_cpf(cls, v: SecretStr) -> SecretStr:
        valor = v.get_secret_value().strip()

        if not valor:
            raise ValueError("CPF não pode ser vazio.")

        if not re.fullmatch(r"\d{3}.\d{3}.\d{3}-\d{2}", valor) and not re.fullmatch(r"\d{11}", valor):
            raise ValueError(
                "CPF deve estar no formato XXX.XXX.XXX-XX ou conter apenas 11 dígitos (XXXXXXXXXXX)."
            )

        return v

    @field_validator("senha")
    @classmethod
    def validar_senha(cls, v: SecretStr) -> SecretStr:
        if not v.get_secret_value().strip():
            raise ValueError("Senha não pode ser vazia.")
        return v

