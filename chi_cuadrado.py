import pandas as pd
import numpy as np
import os

found = './Datos.xlsx'

excel_path = found
try:
    df = pd.read_excel(excel_path)
except Exception as e:
    raise RuntimeError(f"Error al leer el archivo Excel {excel_path}: {e}")

categorias = {
    'Terror': ['Terror'],
    'Comedia': ['Comedia'],
    'Drama': ['Drama']
}

df.columns = [col.strip() for col in df.columns]

def contar_preferencias(df, grupo, generos):
    return df[(df['Grupo de edad'] == grupo) & (df['Género favorito'].isin(generos))].shape[0]

resultados = []
for grupo in ['Adultos', 'Adultos mayores', 'Jóvenes']:
    for genero, lista_generos in categorias.items():
        total = contar_preferencias(df, grupo, lista_generos)
        resultados.append({'Grupo de edad': grupo, 'Género': genero, 'Total': total})

resultados_df = pd.DataFrame(resultados)

observed = resultados_df.pivot(index='Grupo de edad', columns='Género', values='Total').fillna(0).astype(int)

observed_with_totals = observed.copy()
observed_with_totals['Total'] = observed_with_totals.sum(axis=1)
row_totals_obs = observed_with_totals.sum(axis=0).to_frame().T
row_totals_obs.index = ['Total']
observed_with_totals = pd.concat([observed_with_totals, row_totals_obs]).astype(int)

print('--- Tabla de conteos (observados) con totales ---')
print(observed_with_totals)


observed_float = observed.astype(float)

row_totals = observed_float.sum(axis=1)
col_totals = observed_float.sum(axis=0)
grand_total = observed_float.values.sum()

expected = np.outer(row_totals, col_totals) / grand_total
expected_df = pd.DataFrame(expected, index=observed_float.index, columns=observed_float.columns)

with np.errstate(divide='ignore', invalid='ignore'):
    contributions = (observed_float - expected_df) ** 2 / expected_df
    contributions = contributions.replace([np.inf, -np.inf], 0).fillna(0)

chi2_stat = contributions.values.sum()
rows, cols = observed_float.shape
df_degrees = (rows - 1) * (cols - 1)

try:
    from scipy.stats import chi2
    p_value = chi2.sf(chi2_stat, df_degrees)
except ImportError:
    print('scipy no está instalado. El p-valor no se calculará. Puedes instalar scipy con "pip install scipy".')
    p_value = None
except Exception as e:
    print(f'Error al calcular el p-valor: {e}')
    p_value = None

expected_with_totals = expected_df.copy()
expected_with_totals['Total'] = expected_with_totals.sum(axis=1)
row_totals_exp = expected_with_totals.sum(axis=0).to_frame().T
row_totals_exp.index = ['Total']
expected_with_totals = pd.concat([expected_with_totals, row_totals_exp])
expected_with_totals = expected_with_totals.round(2)

contrib_with_totals = contributions.copy()
contrib_with_totals['Total'] = contrib_with_totals.sum(axis=1)
row_totals_contrib = contrib_with_totals.sum(axis=0).to_frame().T
row_totals_contrib.index = ['Total']
contrib_with_totals = pd.concat([contrib_with_totals, row_totals_contrib])
contrib_with_totals = contrib_with_totals.round(4)

print('\n--- Tabla Esperada (bajo H0) con totales ---')
print(expected_with_totals)

print('\n--- Tabla Contribuciones (O - E)^2 / E con totales ---')
print(contrib_with_totals)

print('\n--- Resumen chi-cuadrado ---')
print(f'Estadístico chi2 = {chi2_stat:.4f}')
print(f'Grados de libertad = {df_degrees}')
if p_value is not None:
    print(f'p-valor = {p_value:.6f}')
    alpha = 0.05
    if p_value < alpha:
        print(f'Resultado: p < {alpha} -> Rechazamos la hipótesis nula (hay asociación).')
    else:
        print(f'Resultado: p >= {alpha} -> No rechazamos la hipótesis nula (no hay evidencia de asociación).')
else:
    print('p-valor no calculado (scipy no está disponible o hubo un error). Instala scipy para obtener el p-valor exacto.')

try:
    with pd.ExcelWriter('resultados_preferencias.xlsx') as writer:
        observed_with_totals.to_excel(writer, sheet_name='Observados_con_totales')
        expected_with_totals.to_excel(writer, sheet_name='Esperados_con_totales')
        contrib_with_totals.to_excel(writer, sheet_name='Contribuciones_con_totales')
        resumen = pd.DataFrame({
            'chi2': [chi2_stat],
            'df': [df_degrees],
            'p_value': [p_value if p_value is not None else 'NA']
        })
        resumen.to_excel(writer, sheet_name='Resumen', index=False)
    print('\nLas tres tablas separadas con totales han sido guardadas en "resultados_preferencias.xlsx".')
except PermissionError:
    print('No se pudo guardar el archivo "resultados_preferencias.xlsx". Ciérralo si está abierto e inténtalo de nuevo.')
except Exception as e:
    print(f'Error al guardar el archivo Excel: {e}')
