# Week 3: Advanced Data Analysis and Visualization in Logistics

An end-to-end analytical framework and automated pipeline for exploratory data analysis (EDA), statistical modeling, and multi-modal visualization of supply chain and freight logistics data.

---

## 📌 Project Overview

This repository contains the complete analytical pipeline, dataset, high-resolution visualization suite, and comprehensive executive report for **Week 3: Advanced Data Analysis and Visualization in Logistics**.

The study investigates a **5,000-shipment multi-modal freight network** operating across six major North American logistics hubs (*Chicago, Los Angeles, Dallas, Atlanta, New York, Seattle*), analyzing:
- On-Time In-Full (OTIF) fulfillment dynamics and carrier reliability benchmarks.
- Multi-modal transportation cost drivers and freight consolidation economies of scale.
- Route-specific bottlenecks, lead time variance, and Pareto delay root causes.
- Fourth-quarter (Q4) peak demand seasonality and capacity crunch impacts.

---

## 📁 Repository Structure

```
├── charts/                                      # 300-DPI Publication-Grade Visualizations
│   ├── fig1_cost_vs_weight_by_mode.png          # Cost vs Weight by Mode (Log-Log Scale)
│   ├── fig2_delivery_lead_time_distribution.png # Lead Time Density & Carrier Box Plots
│   ├── fig3_otif_and_carrier_scorecard.png      # Carrier OTIF vs Delay Days Scorecard
│   ├── fig4_delay_root_cause_pareto.png         # Pareto Analysis of Delay Root Causes
│   ├── fig5_correlation_heatmap.png             # Pearson Correlation Matrix Heatmap
│   ├── fig6_route_cost_efficiency_matrix.png    # Route Unit Cost vs OTIF Matrix
│   └── fig7_monthly_volume_and_cost_trend.png   # Monthly Volume & OTIF Seasonality Trend
├── Advanced_Logistics_Data_Analysis_Report.docx # Formatted Executive Word Report
├── logistics_analysis_pipeline.py               # Complete Python Analytics & Report Builder
├── logistics_dataset.csv                        # Simulated Dataset (5,000 Freight Records)
└── README.md                                    # Project Documentation
```

---

## 📊 Key Analytical Findings

| Performance Metric | Observed Baseline | Key Operational Takeaway | Strategic Action |
| :--- | :---: | :--- | :--- |
| **Network OTIF Rate** | **85.3%** | Vendor variance: **92.8%** (*Alpha*) vs **76.1%** (*Nexus*). | Re-allocate volume to top-tier carriers; enforce SLA penalties. |
| **Consolidation Threshold** | **~5,200 kg** | LTL unit costs exceed FTL by 2.3x due to terminal handling. | Establish regional cross-docks in Dallas & Chicago. |
| **Top Delay Root Causes** | **60.3% of total** | Urban Traffic (35.2%) + Severe Weather (25.1%). | Implement dynamic GPS telematics & predictive dispatching. |
| **Q4 Seasonal Shock** | **-9.4% OTIF drop** | Holiday freight volume surges +42%, straining capacity. | Contract dedicated seasonal surge capacity 6 months in advance. |

---

## 🛠️ Installation & Execution

### Prerequisites
- Python 3.9+
- Required libraries:
  ```bash
  pip install pandas numpy matplotlib seaborn python-docx
  ```

### Running the Complete Pipeline
To regenerate the dataset, compute all statistical summaries, generate all 7 visualizations, and compile the final `.docx` report:
```bash
python logistics_analysis_pipeline.py
```

---

## 📄 Deliverables Included
1. **Master DOCX Report**: [`Advanced_Logistics_Data_Analysis_Report.docx`](./Advanced_Logistics_Data_Analysis_Report.docx) containing all sections, tables, embedded figures, justifications, and recommendations.
2. **Dataset**: [`logistics_dataset.csv`](./logistics_dataset.csv) (16 variables across 5,000 records).
3. **Charts Folder**: [`charts/`](./charts/) containing standalone 300-DPI PNGs.
