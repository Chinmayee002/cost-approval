import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { Bar, Doughnut } from "react-chartjs-2";
import "bootstrap/dist/css/bootstrap.min.css";
import "bootstrap/dist/js/bootstrap.bundle.min.js";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

import {
  Chart as ChartJS,
  ArcElement,
  Tooltip,
  Legend,
  CategoryScale,
  LinearScale,
  BarElement,
} from "chart.js";

ChartJS.register(
  ArcElement,
  Tooltip,
  Legend,
  CategoryScale,
  LinearScale,
  BarElement,
);

interface IProps {
  context: any;
}
interface IProduct {
  Id: number;
  ProductID: string;
  ProductName: string;
  Owner: string;
  ApprovalCost: number;
  Status: string;
}
interface IState {
  statusCounts: { [key: string]: number };
  emplist: IProduct[];
}

export default class ProductChart extends React.Component<IProps, IState> {
  constructor(props: IProps) {
    super(props);

    this.state = {
      statusCounts: {},
      emplist: [],
    };
  }

  componentDidMount(): void {
    this.getChartData();
  }

  private getChartData(): void {
    // const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Product Approval Requests')/items?$select=Status`;
    const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Product Approval Requests')/items?$select=Id,ProductID,ProductName,Owner,ApprovalCost,Status`;
    this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((res: Response) => res.json())
      .then((data: { value: IProduct[] }) => {
        const counts: { [key: string]: number } = {};

        data.value.forEach((item: IProduct) => {
          const status = item.Status || "Unknown";
          counts[status] = (counts[status] || 0) + 1;
        });

        this.setState({
          emplist: data.value,
          statusCounts: counts,
        });
      });
  }

  // EXPORT TO EXCEL FUNCTION & CSV FUNCTION
  private exportToExcel = (): void => {
    const { emplist } = this.state;
    console.log("Export Data:", this.state.emplist);
    if (!emplist || emplist.length === 0) {
      alert("No data to export");
      return;
    }

    const exportData = emplist.map((item: IProduct) => ({
      "Product ID": item.ProductID,
      "Product Name": item.ProductName,
      Owner: item.Owner,
      "Approval Cost": item.ApprovalCost,
      Status: item.Status,
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Products");

    const excelBuffer = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
    });

    const blob = new Blob([excelBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    saveAs(blob, "ProductApprovalData.xlsx");
  };

  private exportToCSV = (): void => {
    const { emplist } = this.state;

    if (!emplist || emplist.length === 0) {
      alert("No data to export");
      return;
    }

    const headers = [
      "Product ID",
      "Product Name",
      "Owner",
      "Approval Cost",
      "Status",
    ];

    const rows = emplist.map((item: IProduct) => [
      item.ProductID,
      item.ProductName,
      item.Owner,
      item.ApprovalCost,
      item.Status,
    ]);

    const csvContent = [headers, ...rows]
      .map((row) => row.join(","))
      .join("\n");

    const blob = new Blob([csvContent], {
      type: "text/csv;charset=utf-8;",
    });

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "ProductData.csv";
    link.click();
  };

  render() {
    const labels = Object.keys(this.state.statusCounts);
    const values = labels.map((label) => this.state.statusCounts[label]);
    const options = {
      responsive: true,
      maintainAspectRatio: false,
      layout: {
        padding: 20,
      },
    };
    const chartData = {
      labels: labels,
      datasets: [
        {
          label: "Products",
          data: values,
          backgroundColor: ["#6366F1", "#10B981", "#F43F5E", "#2196F3"],
        },
      ],
    };

    return (
      <div>
        <div className="dropdown ms-2">
          <button
            className="btn btn-success dropdown-toggle"
            type="button"
            data-bs-toggle="dropdown"
          >
            <i className="bi bi-download"></i> Export
          </button>

          <ul className="dropdown-menu">
            <li>
              <button
                className="dropdown-item"
                onClick={this.exportToExcel}
                disabled={this.state.emplist.length === 0}
              >
                Export as Excel (.xlsx)
              </button>
            </li>
            <li>
              <button
                className="dropdown-item"
                onClick={this.exportToCSV}
                disabled={this.state.emplist.length === 0}
              >
                Export as CSV (.csv)
              </button>
            </li>
          </ul>
        </div>
        <h3 className="p-3">Chart Layout</h3>
        <div className="container-fluid">
          <div className="row">
            <div className="col-12 col-md-6">
              <div style={{ height: "450px" }}>
                <Doughnut data={chartData} options={options} />
              </div>{" "}
            </div>
            <div className="col-12 col-md-6">
              <div style={{ height: "450px" }}>
                <Bar data={chartData} options={options} />
              </div>{" "}
            </div>
          </div>
        </div>
        {/* <Doughnut data={chartData} /> */}
      </div>
    );
  }
}
