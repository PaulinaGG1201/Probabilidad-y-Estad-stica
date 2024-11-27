from shiny import App, render, ui
import pandas as pd
import plotly.express as px
import io
import xlsxwriter

# Definir la UI de la aplicación
app_ui = ui.page_fluid(
    ui.h2("Aplicación interactiva de estimadores"),
    ui.input_file("file", "Sube un archivo Excel", accept=[".xlsx"]),
    ui.output_text("preview"),
    ui.output_ui("column"),
    ui.output_text("stats"),
    ui.output_ui("plot"),  # Cambiar de output_plot a output_ui para manejar Plotly como HTML
    ui.download_button("download", "Descargar informe")
)

# Definir la lógica del servidor
def server(input, output, session):
    # Leer y previsualizar el archivo Excel
    @output
    @render.text
    def preview():
        if not input.file():
            return "Por favor, sube un archivo Excel para continuar."
        df = pd.read_excel(input.file()[0]["datapath"])
        return f"Datos cargados correctamente. Columnas disponibles: {', '.join(df.columns)}"

    # Actualizar columnas para selección
    @output
    @render.ui
    def column():
        if not input.file():
            return ui.input_select("column", "Selecciona la columna para análisis:", [], multiple=False)
        df = pd.read_excel(input.file()[0]["datapath"])
        numeric_columns = df.select_dtypes(include="number").columns
        return ui.input_select("column", "Selecciona la columna para análisis:", numeric_columns.tolist(), multiple=False)

    # Mostrar estadísticas descriptivas
    @output
    @render.text
    def stats():
        if not input.file() or not input.column():
            return "Por favor, selecciona una columna para el análisis."
        df = pd.read_excel(input.file()[0]["datapath"])
        column = input.column()
        data = df[column]
        mean = data.mean()
        median = data.median()
        mode = data.mode().iloc[0] if not data.mode().empty else "Sin moda"
        variance = data.var()
        std_dev = data.std()
        return (f"Media: {mean:.2f}\n"
                f"Mediana: {median:.2f}\n"
                f"Moda: {mode}\n"
                f"Varianza: {variance:.2f}\n"
                f"Desviación estándar: {std_dev:.2f}")

    # Generar gráficos interactivos
    @output
    @render.ui
    def plot():
        if not input.file() or not input.column():
            return None
        df = pd.read_excel(input.file()[0]["datapath"])
        column = input.column()
        fig = px.histogram(df, x=column, title=f"Distribución de {column}", nbins=10)
        # Guardar el gráfico como un archivo HTML y renderizarlo
        fig_html = fig.to_html(full_html=False)
        return ui.HTML(fig_html)

    # Generar y descargar informe
    @output
    @render.download
    def download():
        def write_report():
            output = io.BytesIO()
            df = pd.read_excel(input.file()[0]["datapath"])
            column = input.column()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                # Guardar los datos originales
                df.to_excel(writer, index=False, sheet_name="Datos Originales")

                # Agregar estadísticas descriptivas
                stats = {
                    "Estadística": ["Media", "Mediana", "Moda", "Varianza", "Desviación Estándar"],
                    "Valor": [
                        df[column].mean(),
                        df[column].median(),
                        df[column].mode().iloc[0] if not df[column].mode().empty else "Sin moda",
                        df[column].var(),
                        df[column].std(),
                    ]
                }
                stats_df = pd.DataFrame(stats)
                stats_df.to_excel(writer, index=False, sheet_name="Estadísticas")

            output.seek(0)
            return output.getvalue()

        return write_report

# Crear la aplicación
app = App(app_ui, server)





