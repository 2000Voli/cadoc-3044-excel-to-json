# -*- coding: utf-8 -*-
"""
Created on Thu May 14 15:05:14 2026

@author: ov0006
"""


from __future__ import annotations

import json
from pathlib import Path
from dataclasses import dataclass
from typing import Optional, Dict, Any, List
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import numpy as np


EXCEL_PATH = "samples/layout_entrada_3044.xlsx"

OUTPUT_JSON_PATH = "outputs/cadoc_3044.json"

TIMEZONE = "America/Sao_Paulo"


@dataclass
class Pagamento:
    acao: str
    data: str
    valor: Optional[float] = None
    tpMotivo: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        d = {
            "acao": int(self.acao),
            "data": self.data
        }

        # Para exclusão, o manual exige apenas data.
        if str(self.acao) != "2":
            if self.tpMotivo not in [None, "", "nan"]:
                d["tpMotivo"] = str(self.tpMotivo).strip()

            if self.valor is not None and not pd.isna(self.valor):
                d["valor"] = round(float(self.valor), 2)

        return d


@dataclass
class Cessao:
    acao: str
    data: str
    cdCessionario: str
    valor: Optional[float] = None

    def to_dict(self) -> Dict[str, Any]:
        d = {
            "acao": int(self.acao),
            "data": self.data,
            "cdCessionario": self.cdCessionario
        }

        # Para exclusão de cessão, o manual exige data e cdCessionario.
        if str(self.acao) != "2":
            if self.valor is not None and not pd.isna(self.valor):
                d["valor"] = round(float(self.valor), 2)

        return d


@dataclass
class Aquisicao:
    acao: str
    data: str
    cdCedente: Optional[str] = None
    valor: Optional[float] = None

    def to_dict(self) -> Dict[str, Any]:
        d = {
            "acao": int(self.acao),
            "data": self.data
        }

        # Para exclusão de aquisição, o manual exige apenas data.
        if str(self.acao) != "2":
            if self.cdCedente not in [None, "", "nan"]:
                d["cdCedente"] = self.cdCedente

            if self.valor is not None and not pd.isna(self.valor):
                d["valor"] = round(float(self.valor), 2)

        return d


@dataclass
class Operacao:
    acao: str
    ipoc: str
    saldoDevedor: float
    dataSaldoDevedor: str
    atraso: str
    pagamentos: Optional[List[Pagamento]] = None
    cessoes: Optional[List[Cessao]] = None
    aquisicoes: Optional[List[Aquisicao]] = None

    def to_dict(self) -> Dict[str, Any]:
        base = {
            "acao": int(self.acao),
            "ipoc": self.ipoc,
            "saldoDevedor": round(float(self.saldoDevedor), 2),
            "dataSaldoDevedor": self.dataSaldoDevedor,
            "atraso": self.atraso
        }

        if self.pagamentos:
            base["pagamentos"] = [p.to_dict() for p in self.pagamentos]

        if self.cessoes:
            base["cessoes"] = [c.to_dict() for c in self.cessoes]

        if self.aquisicoes:
            base["aquisicoes"] = [a.to_dict() for a in self.aquisicoes]

        return base


def _ensure_cols(df: pd.DataFrame, cols: List[str], sheet_name: str) -> None:
    faltando = [c for c in cols if c not in df.columns]
    if faltando:
        raise ValueError(f"Colunas ausentes na aba '{sheet_name}': {faltando}")


def _read_required_sheet_allow_empty(xls: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    if sheet_name not in xls.sheet_names:
        raise ValueError(f"A aba obrigatória '{sheet_name}' não foi encontrada no Excel.")

    df = pd.read_excel(xls, sheet_name)
    df = df.dropna(how="all")

    return df


def _coerce_date(obj) -> str:
    if pd.isna(obj):
        raise ValueError("Data obrigatória está vazia.")

    if isinstance(obj, (pd.Timestamp, datetime)):
        return obj.strftime("%Y-%m-%d")

    return pd.to_datetime(str(obj), errors="raise").strftime("%Y-%m-%d")


def _now_sp(tz_name: str) -> str:
    return datetime.now(ZoneInfo(tz_name)).strftime("%Y-%m-%d %H:%M:%S")


def _clean_code_8(value, field_name: str, required: bool = True) -> Optional[str]:
    if pd.isna(value):
        if required:
            raise ValueError(f"{field_name} obrigatório está vazio.")
        return None

    raw = str(value).strip()

    if raw == "" or raw.lower() == "nan":
        if required:
            raise ValueError(f"{field_name} obrigatório está vazio.")
        return None

    raw = raw.split(".")[0]
    raw = "".join(ch for ch in raw if ch.isalnum())

    if not raw:
        if required:
            raise ValueError(f"{field_name} inválido.")
        return None

    return raw.zfill(8)[:8]


def _normalize_remessas(remessas: pd.DataFrame) -> Dict[str, Any]:
    _ensure_cols(remessas, ["cnpjIF", "dataHoraRemessa", "envia3050"], "Remessas")

    if remessas.empty:
        raise ValueError("A aba Remessas está vazia.")

    header = remessas.iloc[0].copy()

    cnpjIF = _clean_code_8(header["cnpjIF"], "cnpjIF", required=True)
    envia3050 = str(header["envia3050"]).strip().upper()

    if envia3050 != "N":
        raise ValueError(
            "Este script está configurado para FIDC com envia3050 = 'N'. "
            "Portanto, class3050 não será enviado."
        )

    return {
        "cnpjIF": cnpjIF,
        "envia3050": envia3050
    }


def _normalize_operacoes(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.dropna(how="all")

    _ensure_cols(
        df,
        ["acao", "ipoc", "saldoDevedor", "dataSaldoDevedor", "atraso"],
        "Operacoes"
    )

    if df.empty:
        raise ValueError("A aba Operacoes não pode estar vazia.")

    df["acao"] = df["acao"].astype(str).str.strip()
    df["ipoc"] = df["ipoc"].astype(str).str.strip()
    df["atraso"] = df["atraso"].astype(str).str.upper().str.strip()

    df["dataSaldoDevedor"] = df["dataSaldoDevedor"].apply(_coerce_date)
    df["saldoDevedor"] = pd.to_numeric(df["saldoDevedor"], errors="coerce").fillna(0.0)

    if df["ipoc"].duplicated().any():
        duplicados = df.loc[df["ipoc"].duplicated(), "ipoc"].tolist()
        raise ValueError(f"IPOCs duplicados na aba Operacoes: {duplicados}")

    acoes_validas = {"1", "2"}
    acoes_invalidas = sorted(set(df["acao"]) - acoes_validas)
    if acoes_invalidas:
        raise ValueError(f"Ação inválida na aba Operacoes: {acoes_invalidas}. Use 1 ou 2.")

    atrasos_validos = {"S", "N"}
    atrasos_invalidos = sorted(set(df["atraso"]) - atrasos_validos)
    if atrasos_invalidos:
        raise ValueError(f"Atraso inválido na aba Operacoes: {atrasos_invalidos}. Use S ou N.")

    return df


def _normalize_pagamentos(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    df = df.copy()

    _ensure_cols(
        df,
        ["ipoc", "acao", "data", "valor"],
        "Pagamentos"
    )

    if "tpMotivo" not in df.columns:
        df["tpMotivo"] = None

    df["ipoc"] = df["ipoc"].astype(str).str.strip()
    df["acao"] = df["acao"].astype(str).str.strip()
    df["data"] = df["data"].apply(_coerce_date)
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")

    df["tpMotivo"] = df["tpMotivo"].apply(
        lambda x: None if pd.isna(x) or str(x).strip() == "" else str(x).strip().split(".")[0]
    )

    acoes_validas = {"1", "2", "3"}
    acoes_invalidas = sorted(set(df["acao"]) - acoes_validas)
    if acoes_invalidas:
        raise ValueError(f"Ação inválida na aba Pagamentos: {acoes_invalidas}. Use 1, 2 ou 3.")

    motivos_validos = {"1", "2", None}
    motivos_invalidos = sorted(
        {m for m in df["tpMotivo"].tolist() if m not in motivos_validos}
    )
    if motivos_invalidos:
        raise ValueError(
            f"tpMotivo inválido na aba Pagamentos: {motivos_invalidos}. "
            "Use 1, 2 ou deixe vazio."
        )

    return df


def _normalize_cessoes(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    df = df.copy()

    _ensure_cols(
        df,
        ["ipoc", "acao", "data", "cdCessionario", "valor"],
        "Cessoes"
    )

    df["ipoc"] = df["ipoc"].astype(str).str.strip()
    df["acao"] = df["acao"].astype(str).str.strip()
    df["data"] = df["data"].apply(_coerce_date)

    df["cdCessionario"] = df["cdCessionario"].apply(
        lambda x: _clean_code_8(x, "cdCessionario", required=True)
    )

    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")

    acoes_validas = {"1", "2", "3"}
    acoes_invalidas = sorted(set(df["acao"]) - acoes_validas)
    if acoes_invalidas:
        raise ValueError(f"Ação inválida na aba Cessoes: {acoes_invalidas}. Use 1, 2 ou 3.")

    return df


def _normalize_aquisicoes(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    df = df.copy()

    _ensure_cols(
        df,
        ["ipoc", "acao", "data", "cdCedente", "valor"],
        "Aquisicoes"
    )

    df["ipoc"] = df["ipoc"].astype(str).str.strip()
    df["acao"] = df["acao"].astype(str).str.strip()
    df["data"] = df["data"].apply(_coerce_date)

    df["cdCedente"] = df.apply(
        lambda row: _clean_code_8(
            row["cdCedente"],
            "cdCedente",
            required=str(row["acao"]).strip() != "2"
        ),
        axis=1
    )

    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")

    acoes_validas = {"1", "2", "3"}
    acoes_invalidas = sorted(set(df["acao"]) - acoes_validas)
    if acoes_invalidas:
        raise ValueError(f"Ação inválida na aba Aquisicoes: {acoes_invalidas}. Use 1, 2 ou 3.")

    return df


def _validar_ipocs_eventos(
    operacoes_map: Dict[str, Operacao],
    df_eventos: pd.DataFrame,
    nome_aba: str
) -> None:
    if df_eventos.empty:
        return

    ipocs_operacoes = set(operacoes_map.keys())
    ipocs_eventos = set(df_eventos["ipoc"].astype(str).str.strip())

    faltantes = sorted(ipocs_eventos - ipocs_operacoes)

    if faltantes:
        raise ValueError(
            f"Existem IPOCs na aba {nome_aba} que não existem na aba Operacoes: {faltantes}"
        )


def excel_to_json_3044(
    excel_path: str | Path,
    tz_name: str = TIMEZONE
) -> Dict[str, Any]:

    excel_path = Path(excel_path)

    if not excel_path.exists():
        raise FileNotFoundError(f"Arquivo Excel não encontrado: {excel_path}")

    xls = pd.ExcelFile(excel_path)

    remessas_df = pd.read_excel(xls, "Remessas")
    operacoes_df = pd.read_excel(xls, "Operacoes")

    pagamentos_df = _read_required_sheet_allow_empty(xls, "Pagamentos")
    cessoes_df = _read_required_sheet_allow_empty(xls, "Cessoes")
    aquisicoes_df = _read_required_sheet_allow_empty(xls, "Aquisicoes")

    remessa = _normalize_remessas(remessas_df)

    operacoes_df = _normalize_operacoes(operacoes_df)
    pagamentos_df = _normalize_pagamentos(pagamentos_df)
    cessoes_df = _normalize_cessoes(cessoes_df)
    aquisicoes_df = _normalize_aquisicoes(aquisicoes_df)

    operacoes_map: Dict[str, Operacao] = {}

    for _, row in operacoes_df.iterrows():
        ipoc = row["ipoc"]

        operacoes_map[ipoc] = Operacao(
            acao=row["acao"],
            ipoc=ipoc,
            saldoDevedor=float(row["saldoDevedor"]),
            dataSaldoDevedor=row["dataSaldoDevedor"],
            atraso=row["atraso"],
            pagamentos=[],
            cessoes=[],
            aquisicoes=[]
        )

    _validar_ipocs_eventos(operacoes_map, pagamentos_df, "Pagamentos")
    _validar_ipocs_eventos(operacoes_map, cessoes_df, "Cessoes")
    _validar_ipocs_eventos(operacoes_map, aquisicoes_df, "Aquisicoes")

    if not pagamentos_df.empty:
        pagamentos_df = pagamentos_df.sort_values(["ipoc", "data"], kind="stable")

        for _, row in pagamentos_df.iterrows():
            ipoc = row["ipoc"]

            operacoes_map[ipoc].pagamentos.append(
                Pagamento(
                    acao=row["acao"],
                    data=row["data"],
                    valor=row["valor"],
                    tpMotivo=row["tpMotivo"]
                )
            )

    if not cessoes_df.empty:
        cessoes_df = cessoes_df.sort_values(["ipoc", "data", "cdCessionario"], kind="stable")

        for _, row in cessoes_df.iterrows():
            ipoc = row["ipoc"]

            operacoes_map[ipoc].cessoes.append(
                Cessao(
                    acao=row["acao"],
                    data=row["data"],
                    cdCessionario=row["cdCessionario"],
                    valor=row["valor"]
                )
            )

    if not aquisicoes_df.empty:
        aquisicoes_df = aquisicoes_df.sort_values(["ipoc", "data"], kind="stable")

        for _, row in aquisicoes_df.iterrows():
            ipoc = row["ipoc"]

            operacoes_map[ipoc].aquisicoes.append(
                Aquisicao(
                    acao=row["acao"],
                    data=row["data"],
                    cdCedente=row["cdCedente"],
                    valor=row["valor"]
                )
            )

    operacoes_list = []

    for op in operacoes_map.values():
        if not op.pagamentos:
            op.pagamentos = None

        if not op.cessoes:
            op.cessoes = None

        if not op.aquisicoes:
            op.aquisicoes = None

        operacoes_list.append(op.to_dict())

    json_3044 = {
        "cnpjIF": remessa["cnpjIF"],
        "dataHoraRemessa": _now_sp(tz_name),
        "envia3050": remessa["envia3050"],
        "operacoes": operacoes_list
    }

    return json_3044


def main():
    data = excel_to_json_3044(EXCEL_PATH, tz_name=TIMEZONE)

    out = Path(OUTPUT_JSON_PATH)
    out.parent.mkdir(parents=True, exist_ok=True)

    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"JSON 3044 gerado em: {out}")


if __name__ == "__main__":
    main()