import { createFileRoute } from "@tanstack/react-router";
import { useState } from "react";
import ExcelJS from "exceljs";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Remanejamento Cidade Administrativa" },
      {
        name: "description",
        content:
          "Formulário para registrar remanejamento de dispositivos e exportar para planilha Excel.",
      },
      { property: "og:title", content: "Remanejamento Cidade Administrativa" },
      {
        property: "og:description",
        content:
          "Formulário para registrar remanejamento de dispositivos e exportar para planilha Excel.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary" },
    ],
  }),
  component: Index,
});

const TIPO_DISPOSITIVO = [
  "CISCO IP Phone 7945",
  "CISCO IP Phone 7906",
  "CISCO IP Phone 7965",
  "CISCO IP Phone 7915",
  "COMPUTADOR",
  "NOTEBOOK",
];

const CONDICAO_DESTINO = [
  "VAZIA",
  "TROCA VICE-VERSA",
  "TROCA ALTERNADA",
  "INSTALAÇÃO",
  "REMOÇÃO",
];

const ORIGEM_VAZIA = ["VAZIA", "TROCA VICE-VERSA", "TROCA ALTERNADA"];

const PREDIO_ANDAR = ["PREDIO GERAIS - 4º ANDAR"];


type Dispositivo = {
  tipoDispositivo: string;
  ramal: string;
  patrimonio: string;
  numeroSerie: string;
  estacaoOrigem: string;
  pontoRedeOrigem: string;
  origemVazia: string;
  estacaoDestino: string;
  pontoRedeDestino: string;
  condicaoDestino: string;
};

const linhaVazia = (): Dispositivo => ({
  tipoDispositivo: "",
  ramal: "",
  patrimonio: "",
  numeroSerie: "",
  estacaoOrigem: "",
  pontoRedeOrigem: "",
  origemVazia: "",
  estacaoDestino: "",
  pontoRedeDestino: "",
  condicaoDestino: "",
});

function Index() {
  const [orgao, setOrgao] = useState("");
  const [predio, setPredio] = useState("");
  const [linhas, setLinhas] = useState<Dispositivo[]>([linhaVazia()]);
  const [gerando, setGerando] = useState(false);
  const [erroAberto, setErroAberto] = useState(false);

  const atualizar = (idx: number, campo: keyof Dispositivo, valor: string) => {
    setLinhas((prev) =>
      prev.map((l, i) => (i === idx ? { ...l, [campo]: valor } : l)),
    );
  };

  const addLinha = () => setLinhas((p) => [...p, linhaVazia()]);
  const removerLinha = (idx: number) =>
    setLinhas((p) => p.filter((_, i) => i !== idx));
  const removerTodos = () => setLinhas([]);

  const validar = (): boolean => {
    if (!orgao.trim()) return false;
    if (!predio) return false;
    if (linhas.length === 0) return false;
    const obrig: (keyof Dispositivo)[] = [
      "tipoDispositivo",
      "ramal",
      "patrimonio",
      "numeroSerie",
      "estacaoOrigem",
      "pontoRedeOrigem",
      "origemVazia",
      "estacaoDestino",
      "pontoRedeDestino",
      "condicaoDestino",
    ];
    return linhas.every((l) => obrig.every((k) => String(l[k] ?? "").trim()));
  };

  const gerarExcel = async () => {
    if (!validar()) {
      setErroAberto(true);
      return;
    }


    setGerando(true);
    try {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("Remanejamento");

      // Título
      ws.mergeCells("A1:J1");
      const titulo = ws.getCell("A1");
      titulo.value = "Remanejamento Cidade Administrativa";
      titulo.font = { bold: true, size: 14 };
      titulo.alignment = { horizontal: "center", vertical: "middle" };
      titulo.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFA9D08E" },
      };
      ws.getRow(1).height = 24;

      // Cabeçalho do órgão / prédio
      ws.getCell("A3").value = "ÓRGAO / ENTIDADE";
      ws.getCell("A3").font = { bold: true };
      ws.getCell("A3").alignment = { horizontal: "center", vertical: "middle" };
      ws.mergeCells("B3:J3");
      ws.getCell("B3").value = orgao;
      ws.getCell("B3").alignment = { horizontal: "center", vertical: "middle" };
      ["A3", "B3"].forEach((addr) => {
        ws.getCell(addr).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFA9D08E" },
        };
      });

      ws.getCell("A4").value = "PRÉDIO / ANDAR";
      ws.getCell("A4").font = { bold: true };
      ws.getCell("A4").alignment = { horizontal: "center", vertical: "middle" };
      ws.mergeCells("B4:J4");
      ws.getCell("B4").value = predio || "PRÉDIO MINAS -  º ANDAR";
      ws.getCell("B4").alignment = { horizontal: "center", vertical: "middle" };
      ["A4", "B4"].forEach((addr) => {
        ws.getCell(addr).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFA9D08E" },
        };
      });


      // Cabeçalhos das colunas
      const headers = [
        "TIPO DISPOSITIVO",
        "N° RAMAL",
        "PATRIMÔNIO",
        "NÚMERO SÉRIE",
        "NÚMERO DA ESTAÇÃO DE ORIGEM",
        "N° PONTO REDE ESTAÇÃO ORIGEM",
        "ESTAÇÃO DE ORIGEM FICARÁ VAZIA?",
        "NÚMERO DA ESTAÇÃO DE DESTINO",
        "N° PONTO REDE ESTAÇÃO DESTINO",
        "CONDIÇÃO DA ESTAÇÃO DE DESTINO",
      ];
      const headerRow = ws.getRow(5);
      headers.forEach((h, i) => {
        const c = headerRow.getCell(i + 1);
        c.value = h;
        c.font = { bold: true, size: 10 };
        c.alignment = { wrapText: true, horizontal: "center", vertical: "middle" };
        c.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFA9D08E" },
        };
        c.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });
      headerRow.height = 45;

      // Linhas de dados
      linhas.forEach((d, i) => {
        const row = ws.getRow(6 + i);
        const vals = [
          d.tipoDispositivo,
          d.ramal,
          d.patrimonio,
          d.numeroSerie,
          d.estacaoOrigem,
          d.pontoRedeOrigem,
          d.origemVazia,
          d.estacaoDestino,
          d.pontoRedeDestino,
          d.condicaoDestino,
        ];
        vals.forEach((v, j) => {
          const c = row.getCell(j + 1);
          c.value = v;
          c.alignment = { wrapText: true, vertical: "middle" };
          c.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        });
      });

      // Larguras
      const widths = [22, 10, 14, 16, 18, 18, 20, 18, 18, 20];
      widths.forEach((w, i) => {
        ws.getColumn(i + 1).width = w;
      });

      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const nomeArquivo = `remanejamento-${new Date()
        .toISOString()
        .slice(0, 10)}.xlsx`;
      a.download = nomeArquivo;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } finally {
      setGerando(false);
    }
  };

  return (
    <div className="min-h-screen bg-background p-6">
      <div className="mx-auto max-w-6xl">
        <header className="mb-6">
          <h1 className="text-3xl font-bold tracking-tight">
            Remanejamento Cidade Administrativa
          </h1>
          <p className="mt-1 text-sm text-muted-foreground">
            Preencha os dados abaixo e gere a planilha Excel pronta para envio.
          </p>
          <p className="mt-2 text-sm">
            Campos marcados com{" "}
            <span className="font-semibold text-destructive">*</span> são de
            preenchimento obrigatório.
          </p>
        </header>

        <section className="mb-6 grid gap-4 rounded-lg border bg-card p-4 sm:grid-cols-2">
          <div>
            <label className="mb-1 block text-sm font-medium">
              Órgão / Entidade <span className="text-destructive">*</span>
            </label>
            <input
              type="text"
              value={orgao}
              onChange={(e) => setOrgao(e.target.value)}
              className="w-full rounded-md border bg-background px-3 py-2 text-sm"
              placeholder="IPSEMG"
            />
          </div>
          <div>
            <label className="mb-1 block text-sm font-medium">
              Prédio / Andar <span className="text-destructive">*</span>
            </label>
            <select
              value={predio}
              onChange={(e) => setPredio(e.target.value)}
              className="w-full rounded-md border bg-background px-3 py-2 text-sm"
            >
              <option value="">Selecione…</option>
              {PREDIO_ANDAR.map((p) => (
                <option key={p} value={p}>
                  {p}
                </option>
              ))}
            </select>
          </div>
        </section>

        <section className="mb-4 space-y-4">
          {linhas.map((l, idx) => (
            <div
              key={idx}
              className="rounded-lg border bg-card p-4"
            >
              <div className="mb-3 flex items-center justify-between">
                <h2 className="font-semibold">Dispositivo {idx + 1}</h2>
                <button
                  type="button"
                  onClick={() => removerLinha(idx)}
                  className="text-sm text-destructive hover:underline"
                >
                  Remover
                </button>
              </div>

              <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
                <Campo label="Tipo de dispositivo" required>
                  <select
                    value={l.tipoDispositivo}
                    onChange={(e) =>
                      atualizar(idx, "tipoDispositivo", e.target.value)
                    }
                    className="w-full rounded-md border bg-background px-3 py-2 text-sm"
                  >
                    <option value="">Selecione…</option>
                    {TIPO_DISPOSITIVO.map((t) => (
                      <option key={t} value={t}>
                        {t}
                      </option>
                    ))}
                  </select>
                </Campo>

                <Campo label="Nº Ramal" required>
                  <Input
                    value={l.ramal}
                    onChange={(v) => atualizar(idx, "ramal", v)}
                  />
                </Campo>
                <Campo label="Patrimônio" required>
                  <Input
                    value={l.patrimonio}
                    onChange={(v) => atualizar(idx, "patrimonio", v)}
                  />
                </Campo>
                <Campo label="Número de série" required>
                  <Input
                    value={l.numeroSerie}
                    onChange={(v) => atualizar(idx, "numeroSerie", v)}
                  />
                </Campo>
                <Campo label="Nº estação de origem" required>
                  <Input
                    value={l.estacaoOrigem}
                    onChange={(v) => atualizar(idx, "estacaoOrigem", v)}
                  />
                </Campo>
                <Campo label="Nº ponto de rede (origem)" required>
                  <Input
                    value={l.pontoRedeOrigem}
                    onChange={(v) => atualizar(idx, "pontoRedeOrigem", v)}
                  />
                </Campo>
                <Campo label="Estação de origem ficará vazia?" required>
                  <select
                    value={l.origemVazia}
                    onChange={(e) =>
                      atualizar(idx, "origemVazia", e.target.value)
                    }
                    className="w-full rounded-md border bg-background px-3 py-2 text-sm"
                  >
                    <option value="">Selecione…</option>
                    {ORIGEM_VAZIA.map((t) => (
                      <option key={t} value={t}>
                        {t}
                      </option>
                    ))}
                  </select>
                </Campo>
                <Campo label="Nº estação de destino" required>
                  <Input
                    value={l.estacaoDestino}
                    onChange={(v) => atualizar(idx, "estacaoDestino", v)}
                  />
                </Campo>
                <Campo label="Nº ponto de rede (destino)" required>
                  <Input
                    value={l.pontoRedeDestino}
                    onChange={(v) => atualizar(idx, "pontoRedeDestino", v)}
                  />
                </Campo>
                <Campo label="Condição da estação de destino" required>
                  <select
                    value={l.condicaoDestino}
                    onChange={(e) =>
                      atualizar(idx, "condicaoDestino", e.target.value)
                    }
                    className="w-full rounded-md border bg-background px-3 py-2 text-sm"
                  >
                    <option value="">Selecione…</option>
                    {CONDICAO_DESTINO.map((t) => (
                      <option key={t} value={t}>
                        {t}
                      </option>
                    ))}
                  </select>
                </Campo>
              </div>
            </div>
          ))}
        </section>

        <div className="flex flex-wrap gap-3">
          <button
            type="button"
            onClick={addLinha}
            className="rounded-md border px-4 py-2 text-sm font-medium hover:bg-muted"
          >
            + Adicionar dispositivo
          </button>
          <button
            type="button"
            onClick={removerTodos}
            disabled={linhas.length === 0}
            className="rounded-md border border-destructive/40 px-4 py-2 text-sm font-medium text-destructive hover:bg-destructive/10 disabled:opacity-40"
          >
            Remover todos
          </button>
          <button
            type="button"
            onClick={gerarExcel}
            disabled={gerando}
            className="rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-60"
          >
            {gerando ? "Gerando…" : "Salvar planilha Excel"}
          </button>
        </div>

        <footer className="mt-10 border-t pt-4 text-center text-xs text-muted-foreground">
          <p>Desenvolvido por: Wesley Fernandes</p>

        </footer>
      </div>

      {erroAberto && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4"
          onClick={() => setErroAberto(false)}
        >
          <div
            role="alertdialog"
            className="w-full max-w-sm rounded-lg border bg-card p-6 text-center shadow-lg"
            onClick={(e) => e.stopPropagation()}
          >
            <h2 className="mb-2 text-lg font-semibold text-destructive">
              Atenção
            </h2>
            <p className="mb-4 text-sm">Preencha os campos obrigatórios.</p>
            <button
              type="button"
              onClick={() => setErroAberto(false)}
              className="rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:opacity-90"
            >
              OK
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

function Campo({
  label,
  required,
  children,
}: {
  label: string;
  required?: boolean;
  children: React.ReactNode;
}) {
  return (
    <div>
      <label className="mb-1 block text-xs font-medium text-muted-foreground">
        {label}
        {required && <span className="ml-0.5 text-destructive">*</span>}
      </label>
      {children}
    </div>
  );
}

function Input({
  value,
  onChange,
}: {
  value: string;
  onChange: (v: string) => void;
}) {
  return (
    <input
      type="text"
      value={value}
      onChange={(e) => onChange(e.target.value)}
      className="w-full rounded-md border bg-background px-3 py-2 text-sm"
    />
  );
}
