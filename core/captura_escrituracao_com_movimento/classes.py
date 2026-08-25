from pydantic import BaseModel, field_validator
from enum import Enum
from datetime import date
from typing import Any
import re


# ========================== Enums ==========================
class EnumTipoMovimentoEmpresa(Enum):
    SEM_MOVIMENTO = ["SM-", "SM -", "S/M", "SEM MOVIMENTO", "SEMMOVIMENTO", "INATIVA", "NÃO FAZ FISCAL", "NÃOFAZFISCAL"]

# ========================== Classes ==========================

class Empresa(BaseModel):
    """Representa um registro (empresa) na planilha do ISS"""

    codigo: int
    nome: str
    cnpj: str
    responsavel: str   # Campo utilizado apenas p/ verificar se a empresa é sem movimento
    socio_ativo_reinf: str
    responsavel_separacao: str
    mensagens: list[MensagemISSFortaleza] = []

    @field_validator("cnpj")
    @classmethod
    def validar_cnpj(cls, valor: str) -> str:
        cnpj = re.sub(r"\D", "", valor)

        # Valida a quantidade de dígitos.
        if len(cnpj) != 14:
            raise ValueError(f"CNPJ deve conter exatamente 14 dígitos, ´{cnpj}´ contém {len(cnpj)}.")

        # Valida se todos os dígitos são iguais (ex: 00000000000000, 11111111111111, etc.)
        if cnpj == cnpj[0] * 14:
            raise ValueError(f"CNPJ inválido: {cnpj}")

        return cnpj

    def eh_sem_movimento(self) -> bool:
        """ Determina se a empresa é sem movimento
        :return: True se for sem movimento ou False caso contrário.
        """
        for mov in EnumTipoMovimentoEmpresa.SEM_MOVIMENTO.value:
            if (mov in self.responsavel) or (mov in self.responsavel_separacao) or (mov in self.socio_ativo_reinf):
                return True
        return False

    def __str__(self):
        return f"{self.nome} (Código: {self.codigo}, CNPJ: {self.cnpj})"


class MensagemISSFortaleza(BaseModel):
    """Representa uma mensagem da caixa de entrada do ISS Fortaleza"""
    title: str
    date: date
    content: str


# ========================== Exceptions ==========================

class ThreadStoppedException(Exception):
    pass