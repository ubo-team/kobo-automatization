import streamlit as st
import ubo_ui
import pandas as pd
import numpy as np
from collections import defaultdict
import pymc as pm
import arviz as az

st.set_page_config(layout="wide")

with ubo_ui.card("Të dhënat dhe modeli", step=1):
    uploaded_file = st.file_uploader(
        "Ngarkoni dokumentin CSV me të dhënat MaxDiff", 
        type="csv", 
        help="Dokumenti duhet të ketë kolonat: Attribute 1–5, Best, Worst (dhe Response ID për modelin HB)"
    )

    model_choice = st.selectbox("Modeli i analizës", [
        "Numërimi i thjeshtë", 
        "Analiza Bayesiane hierarkike (HB)"
    ])
    run_analysis = st.button("Fillo analizën", type="primary", disabled=not uploaded_file)

if uploaded_file and run_analysis:
    with ubo_ui.card("Rezultatet", step=2):
        df = pd.read_csv(uploaded_file)

        if model_choice == "Numërimi i thjeshtë":
            st.subheader("Rezultatet: numërimi i thjeshtë")
            attribute_cols = [f"Attribute {i}" for i in range(1, 6)]
            best_counts = defaultdict(int)
            worst_counts = defaultdict(int)
            appearance_counts = defaultdict(int)
            warnings = []

            for _, row in df.iterrows():
                attributes_shown = [row[col] for col in attribute_cols]
                best = row["Best"]
                worst = row["Worst"]

                for attr in attributes_shown:
                    appearance_counts[attr] += 1

                if best in attributes_shown:
                    best_counts[best] += 1
                else:
                    warnings.append(f"Zgjedhja më e mirë '{best}' nuk gjendet te {attributes_shown}")

                if worst in attributes_shown:
                    worst_counts[worst] += 1
                else:
                    warnings.append(f"Zgjedhja më e keqe '{worst}' nuk gjendet te {attributes_shown}")

            all_attrs = sorted(set(appearance_counts.keys()))
            results = []

            for attr in all_attrs:
                best = best_counts.get(attr, 0)
                worst = worst_counts.get(attr, 0)
                appeared = appearance_counts[attr]
                score = (best - worst) / appeared if appeared > 0 else 0
                results.append({
                    "Atributi": attr,
                    "Herë si më i miri": best,
                    "Herë si më i keqi": worst,
                    "Herë i shfaqur": appeared,
                    "Pikët (numërimi i thjeshtë)": round(score, 3)
                })

            results_df = pd.DataFrame(results).sort_values(by="Pikët (numërimi i thjeshtë)", ascending=False)
            st.dataframe(results_df, use_container_width=True)

            if warnings:
                with st.expander("Paralajmërime"):
                    for w in warnings:
                        st.write(w)

            csv = results_df.to_csv(index=False)
            st.download_button("Shkarko rezultatet (CSV)", csv, file_name="simple_count_analysis_results.csv")

        elif model_choice == "Analiza Bayesiane hierarkike (HB)":
            st.subheader("Rezultatet: analiza Bayesiane hierarkike (HB)")

            attribute_cols = [f"Attribute {i}" for i in range(1, 6)]
            attributes = sorted(set(df[attribute_cols].values.flatten()))
            attr_index = {attr: i for i, attr in enumerate(attributes)}
            n_attrs = len(attributes)

            pairwise_data = []
            respondent_ids = []

            for _, row in df.iterrows():
                respondent = row["Response ID"]
                attrs = [row[col] for col in attribute_cols]
                best = row["Best"]
                worst = row["Worst"]

                if best in attrs and worst in attrs:
                    pairwise_data.append((attr_index[best], attr_index[worst]))
                    respondent_ids.append(respondent)

            if not pairwise_data:
                st.error("Nuk u gjetën të dhëna të vlefshme (Best/Worst) për modelin HB.")
            else:
                respondents = sorted(set(respondent_ids))
                respondent_map = {resp: i for i, resp in enumerate(respondents)}

                best_ids = np.array([b for b, w in pairwise_data])
                worst_ids = np.array([w for b, w in pairwise_data])
                resp_ids = np.array([respondent_map[r] for r in respondent_ids])
                n_resp = len(respondents)

                with st.spinner("Po trajnohet modeli Bayesian..."):
                    with pm.Model() as model:
                        mu = pm.Normal("mu", mu=0, sigma=1, shape=n_attrs)
                        sigma = pm.HalfNormal("sigma", sigma=1)
                        utilities = pm.Normal("utilities", mu=mu, sigma=sigma, shape=(n_resp, n_attrs))
                        u_diff = utilities[resp_ids, best_ids] - utilities[resp_ids, worst_ids]
                        observed_data = np.ones(len(best_ids), dtype=np.int8)
                        pm.Bernoulli("obs", logit_p=u_diff, observed=observed_data)
                        trace = pm.sample(1000, tune=2000, target_accept=0.95, chains=4, return_inferencedata=True)

                summary_df = az.summary(trace, var_names=["mu"])
                summary_df.index = [f"mu[{i}]" for i in range(len(summary_df))]
                summary_df["Atributi"] = [attributes[i] for i in range(len(attributes))]
                summary_df = summary_df.reset_index(drop=True)
            
                # Calculate Relative Importance (0–100)
                min_util = summary_df["mean"].min()
                max_util = summary_df["mean"].max()
                summary_df["Rëndësia relative (0–100)"] = ((summary_df["mean"] - min_util) / (max_util - min_util) * 100).round(1)

                # Sort by importance descending
                summary_df = summary_df.sort_values(by="Rëndësia relative (0–100)", ascending=False)
                st.dataframe(summary_df, use_container_width=True)

                csv = summary_df.to_csv(index=False)
                st.download_button("Shkarko rezultatet (CSV)", csv, file_name="hb_analysis_results.csv")
