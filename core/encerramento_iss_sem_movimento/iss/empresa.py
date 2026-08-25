from pathlib import Path
from datetime import date
from pydantic import BaseModel, PositiveInt, field_validator, Field
import re


def agrupar_empresas_por_problema(
        empresas: list[EmpresaSemMovimentoISSFortaleza]
) -> dict[str, list[EmpresaSemMovimentoISSFortaleza]]:
    problems: dict[str, list[EmpresaSemMovimentoISSFortaleza]] = {}
    for emp in empresas:
        for p in emp.problemas:
            p_string = p.upper().strip()
            if p_string not in problems:
                problems[p_string] = [emp]
            else:
                problems[p_string].append(emp)

    return problems


class EmpresaSemMovimentoISSFortaleza(BaseModel):
    """Representa um registro (empresa) na planilha do ISS"""

    codigo: PositiveInt
    nome: str
    cnpj: str
    responsavel: str
    municipio: str
    linha_planilha_fiscal: int
    cnpj_m: str = Field(default_factory=str)  # CNPJ com máscara
    problemas: list[str] = []
    mensagens: list[MensagemISS] = []
    dt_primeiro_encerramento: date | None = Field(default=None)
    caminho_certificado: Path | None = Field(default=None)

    def __str__(self):
        return f"{self.nome} (Código: {self.codigo}, CNPJ: {self.cnpj})"
    
    @field_validator("cnpj")
    @classmethod
    def validar_cnpj(cls, valor: str) -> str:
        cnpj = re.sub(r"\D", "", valor)  # Remove tudo que não for número

        # Valida a quantidade de dígitos.
        if len(cnpj) != 14:
            raise ValueError(f"CNPJ deve conter exatamente 14 dígitos, ´{cnpj}´ contém {len(cnpj)}.")

        # Valida se todos os dígitos são iguais (ex: 00000000000000, 11111111111111, etc.)
        if cnpj == cnpj[0] * 14:
            raise ValueError(f"CNPJ inválido: {cnpj}")

        return cnpj
    
    @field_validator("nome")
    @classmethod
    def validar_nome(cls, valor: str) -> str:
        nome = valor.strip()
        if nome == "":
            raise ValueError(f"Nome inválido ou ausente: `{nome}`")
        return nome.upper()
    
    @field_validator("responsavel")
    @classmethod
    def validar_responsavel(cls, valor: str) -> str:
        responsavel = valor.strip().upper()

        # Validar se não é string vazia
        resp = re.sub(r"\s", "", responsavel)
        if resp == "":
            raise ValueError(f"Responsável inválido ou ausente: `{resp}`")

        # Validar se não é sem movimento
        filtros: list[str] = ["SM -", "SEM MOVIMENTO -", "INATIVA -", "INATIVAS -"]
        if not any(map(lambda x: x in responsavel, filtros)):
            raise ValueError("Empresa não é sem movimento")


        return responsavel
    
    @field_validator("municipio")
    @classmethod
    def validar_municipio(cls, valor: str) -> str:
        municipio = valor.strip().upper()
        if municipio not in ["FORTALEZA", "FOR", "FORT"]:
            raise ValueError(f"Empresa de município `${municipio}` não é de Fortaleza")
        return municipio
    

class MensagemISS(BaseModel):
    """Representa uma mensagem da caixa de entrada do ISS Fortaleza"""
    title: str
    date: date
    content: str
