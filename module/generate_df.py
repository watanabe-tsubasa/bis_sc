import polars as pl

def generate_df(headers: list, body: list):
  df = pl.DataFrame(data=body, schema=headers)
  selected_cols = [col for col in df.columns if '売上高' in col and '前日' not in col]
  df_selected = df.select(selected_cols)
  for col in df_selected.columns:
    df_selected = (df_selected
                  .with_columns(
                    pl.col(col)
                    .str.replace(" ", "0")
                    .str.replace("\xa0", "0")
                    .str.replace(",", "")
                    .cast(pl.Float64)
                    )
                  )
  base_cols = ['店舗CD', '店舗', 'グループCD', 'グループ']
  df_base = df.select(base_cols)
  
  return (df_base, df_selected)