import streamlit as st
import pandas as pd
import numpy as np
from collections import defaultdict
import pymc as pm
import arviz as az

st.set_page_config(page_title="MaxDiff Analyzer", layout="wide")
st.title("MaxDiff Analysis Tool")

uploaded_file = st.file_uploader("Upload your MaxDiff CSV file", type="csv", help="File must include: Attribute 1–5, Best, Worst")

model_choice = st.selectbox("Choose analysis model", [
    "Simple Count Analysis", 
    "Hierarchical Bayesian (HB) Analysis"
])

if uploaded_file and st.button("Run Analysis"):
    df = pd.read_csv(uploaded_file)

    if model_choice == "Simple Count Analysis":
        st.subheader("Results: Simple Count Analysis")
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
                warnings.append(f"Warning: Best item '{best}' not found in {attributes_shown}")

            if worst in attributes_shown:
                worst_counts[worst] += 1
            else:
                warnings.append(f"Warning: Worst item '{worst}' not found in {attributes_shown}")

        all_attrs = sorted(set(appearance_counts.keys()))
        results = []

        for attr in all_attrs:
            best = best_counts.get(attr, 0)
            worst = worst_counts.get(attr, 0)
            appeared = appearance_counts[attr]
            score = (best - worst) / appeared if appeared > 0 else 0
            results.append({
                "Attribute": attr,
                "Best Count": best,
                "Worst Count": worst,
                "Times Shown": appeared,
                "Score (Simple Count Analysis)": round(score, 3)
            })

        results_df = pd.DataFrame(results).sort_values(by="Rezultati (Analizë e thjeshtë)", ascending=False)
        st.dataframe(results_df, use_container_width=True)

        if warnings:
            with st.expander("Warnings"):
                for w in warnings:
                    st.write(w)

        csv = results_df.to_csv(index=False)
        st.download_button("Shkarko rezultatet (CSV)", csv, file_name="rezultatet_analize_thjeshte.csv")

    elif model_choice == "Hierarchical Bayesian (HB) Analysis":
        st.subheader("Results: Hierarchical Bayesian (HB) Analysis")

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
            st.error("No valid pairwise data found for HB model.")
        else:
            respondents = sorted(set(respondent_ids))
            respondent_map = {resp: i for i, resp in enumerate(respondents)}

            best_ids = np.array([b for b, w in pairwise_data])
            worst_ids = np.array([w for b, w in pairwise_data])
            resp_ids = np.array([respondent_map[r] for r in respondent_ids])
            n_resp = len(respondents)

            with st.spinner("Training Bayesian model..."):
                with pm.Model() as model:
                    mu = pm.Normal("mu", mu=0, sigma=1, shape=n_attrs)
                    sigma = pm.HalfNormal("sigma", sigma=1)
                    utilities = pm.Normal("utilities", mu=mu, sigma=sigma, shape=(n_resp, n_attrs))
                    u_diff = utilities[resp_ids, best_ids] - utilities[resp_ids, worst_ids]
                    pm.Bernoulli("obs", logit_p=u_diff, observed=np.ones_like(u_diff))
                    trace = pm.sample(1000, tune=1000, target_accept=0.9, chains=2, return_inferencedata=True)

            summary_df = az.summary(trace, var_names=["mu"])
            summary_df.index = [f"mu[{i}]" for i in range(len(summary_df))]
            summary_df["Attribute"] = [attributes[i] for i in range(len(attributes))]
            summary_df = summary_df.reset_index(drop=True)
            st.dataframe(summary_df, use_container_width=True)

            csv = summary_df.to_csv(index=False)
            st.download_button("Shkarko rezultatet (CSV)", csv, file_name="rezultatet_HB.csv")
