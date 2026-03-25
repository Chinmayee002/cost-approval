import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { IHelloWorldProps } from "./IHelloWorldProps";
import styles from "./HelloWorld.module.scss";

interface IProduct {
  Id: number;
  ProductID: string;
  ProductName: string;
  Owner: string;
  ApprovalCost: number;
  Status: string;
}
interface IHelloWorldState {
  emplist: IProduct[];
  allEmplist: IProduct[];
  isLoading: boolean;
  error: string;
  sortColumn: keyof IProduct | "";
  sortDirection: "asc" | "desc";
  searchText: string;

  showDeleteModal: boolean;
  selectedItemId: number | null;

  showEditModal: boolean;
  editItem: IProduct | null;
  isEditMode: boolean;

  showToast: boolean;
  toastMessage: string;

  currentPage: number;
  itemsPerPage: number;

  selectedStatus: string;
}

export default class HelloWorld extends React.Component<
  IHelloWorldProps,
  IHelloWorldState
> {
  constructor(props: IHelloWorldProps) {
    super(props);

    this.state = {
      emplist: [],
      allEmplist: [],
      isLoading: true,
      error: "",
      sortColumn: "",
      sortDirection: "asc",
      searchText: "",

      showDeleteModal: false,
      selectedItemId: null,

      showEditModal: false,
      editItem: null,
      isEditMode: false,

      showToast: false,
      toastMessage: "",

      currentPage: 1,
      itemsPerPage: 5,

      selectedStatus: "All",
    };
  }
  private toastTimer?: number;

  componentDidMount(): void {
    this.getItems();
  }
  private startToastTimer(): void {
    if (this.toastTimer) {
      clearTimeout(this.toastTimer);
    }

    this.toastTimer = window.setTimeout(() => {
      this.setState({ showToast: false });
    }, 1000);
  }
  private onTabChange = (status: string): void => {
    let filtered = this.state.allEmplist;

    if (status !== "All") {
      filtered = filtered.filter((item) => item.Status === status);
    }

    this.setState({
      selectedStatus: status,
      emplist: filtered,
      currentPage: 1,
    });
  };
  /**
   * GET LIST ITEMS
   */
  private getItems(): void {
    const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Product Approval Requests')/items?$select=Id,ProductID,ProductName,Owner,ApprovalCost,Status`;
    this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((res) => res.json())
      .then((data) => {
        console.log("API Response:", data); // Debug log
        // Map the data to extract Owner name
        const processedData = data.value;
        this.setState({
          emplist: processedData,
          allEmplist: processedData,
          isLoading: false,
        });
      })
      .catch((error) => {
        console.error("Error loading items:", error);
        this.setState({ isLoading: false, error: "Failed to load items" });
      });
  }

  /**
   * SEARCH
   */
  private onSearch = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const searchText = event.target.value.toLowerCase();
    const { selectedStatus, allEmplist } = this.state;

    let filteredData = allEmplist;

    // Status filter
    if (selectedStatus !== "All") {
      filteredData = filteredData.filter(
        (item) => item.Status === selectedStatus,
      );
    }

    // Search filter
    filteredData = filteredData.filter(
      (item) =>
        (item.ProductID &&
          item.ProductID.toLowerCase().indexOf(searchText) !== -1) ||
        (item.Owner && item.Owner.toLowerCase().indexOf(searchText) !== -1) ||
        (item.ProductName &&
          item.ProductName.toLowerCase().indexOf(searchText) !== -1) ||
        (item.Status && item.Status.toLowerCase().indexOf(searchText) !== -1) ||
        (item.ApprovalCost &&
          item.ApprovalCost.toString().indexOf(searchText) !== -1),
    );

    this.setState({
      searchText,
      emplist: filteredData,
      currentPage: 1,
    });
  };

  /**
   * SORT
   */
  private onSort = (column: keyof IProduct): void => {
    const { sortColumn, sortDirection, emplist } = this.state;

    const direction =
      sortColumn === column && sortDirection === "asc" ? "desc" : "asc";

    const sorted = [...emplist].sort((a, b) => {
      if (column === "ApprovalCost") {
        return direction === "asc"
          ? (a.ApprovalCost as number) - (b.ApprovalCost as number)
          : (b.ApprovalCost as number) - (a.ApprovalCost as number);
      }

      return direction === "asc"
        ? String(a[column]).localeCompare(String(b[column]))
        : String(b[column]).localeCompare(String(a[column]));
    });

    this.setState({
      emplist: sorted,
      sortColumn: column,
      sortDirection: direction,
    });
  };

  private renderSortIcon(column: keyof IProduct) {
    if (this.state.sortColumn !== column) return "⇅";

    return this.state.sortDirection === "asc" ? "▲" : "▼";
  }

  /**
   * DELETE MODAL
   */
  private openDeleteModal = (id: number) => {
    this.setState({
      showDeleteModal: true,
      selectedItemId: id,
    });
  };

  private closeDeleteModal = () => {
    this.setState({
      showDeleteModal: false,
    });
  };

  private confirmDelete = async () => {
    const id = this.state.selectedItemId;

    if (!id) return;

    const url =
      `${this.props.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('Product Approval Requests')/items(${id})`;
    await this.props.context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          "X-HTTP-Method": "DELETE",
          "IF-MATCH": "*",
        },
      },
    );

    this.setState((prev) => ({
      emplist: prev.emplist.filter((i) => i.Id !== id),
      allEmplist: prev.allEmplist.filter((i) => i.Id !== id),
      showDeleteModal: false,
      showToast: true,
      toastMessage: "Product deleted successfully",
    }));
    this.startToastTimer();
  };
  private addItem = async (): Promise<void> => {
    const { editItem } = this.state;
    if (!editItem) return;

    const webUrl = this.props.context.pageContext.web.absoluteUrl.trim();
    const listName = "Product Approval Requests";

    const url = `${webUrl}/_api/web/lists/getbytitle('${listName}')/items`;

    const body = JSON.stringify({
      ProductID: editItem.ProductID,
      ProductName: editItem.ProductName,
      ApprovalCost: editItem.ApprovalCost,
      Status: editItem.Status,
    });

    try {
      const response = await this.props.context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata.metadata=none",
            "Content-Type": "application/json;odata.metadata=none",
          },
          body,
        },
      );

      if (!response.ok) {
        const errText = await response.text();
        console.error("Add API error response:", errText);
        throw new Error(`Add failed: ${response.status}`);
      }

      const newItem = await response.json();

      this.setState((prev) => ({
        emplist: [newItem, ...prev.emplist],
        allEmplist: [newItem, ...prev.allEmplist],
        showEditModal: false,
        editItem: null,
        showToast: true,
        toastMessage: "Product added successfully",
      }));

      this.startToastTimer();
    } catch (err) {
      console.error("Add error:", err);
      alert("Error adding employee");
    }
  };
  /**
   * EDIT MODAL
   */
  private openEditModal = (item: IProduct) => {
    this.setState({
      showEditModal: true,
      editItem: { ...item },
      isEditMode: true,
    });
  };

  private openAddModal = (): void => {
    this.setState({
      showEditModal: true,
      isEditMode: false,
      editItem: {
        Id: 0,
        ProductID: "",
        ProductName: "",
        ApprovalCost: 0,
        Status: "Pending",
        Owner: "",
      },
    });
  };

  private closeEditModal = (): void => {
    this.setState({
      showEditModal: false,
      editItem: null,
    });
  };

  /**
   * RENDER
   */
  private updateItem = async (): Promise<void> => {
    const { editItem } = this.state;

    if (!editItem) return;

    const url =
      `${this.props.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('Product Approval Requests')/items(${editItem.Id})`;

    const body = JSON.stringify({
      ProductName: editItem.ProductName,
      ApprovalCost: editItem.ApprovalCost,
      Status: editItem.Status,
      Owner: editItem.Owner,
    });

    try {
      const response = await this.props.context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json",
            "Content-Type": "application/json",
            "X-HTTP-Method": "MERGE",
            "IF-MATCH": "*",
          },
          body,
        },
      );
      if (!response.ok) {
        throw new Error(`Update failed: ${response.status}`);
      }
    } catch (err) {
      console.error("Update error:", err);
      this.setState({
        showToast: true,
        toastMessage: "Error updating product",
      });
      this.startToastTimer();
      return;
    }

    this.setState((prev) => ({
      emplist: prev.emplist.map((i) => (i.Id === editItem.Id ? editItem : i)),
      allEmplist: prev.allEmplist.map((i) =>
        i.Id === editItem.Id ? editItem : i,
      ),
      showEditModal: false,
      showToast: true,
      toastMessage: "Product updated successfully",
    }));
  };
  private goToPage = (page: number): void => {
    this.setState({ currentPage: page });
  };

  private nextPage = (): void => {
    const totalPages = Math.ceil(
      this.state.emplist.length / this.state.itemsPerPage,
    );

    if (this.state.currentPage < totalPages) {
      this.setState((prev) => ({
        currentPage: prev.currentPage + 1,
      }));
    }
  };

  private prevPage = (): void => {
    if (this.state.currentPage > 1) {
      this.setState((prev) => ({
        currentPage: prev.currentPage - 1,
      }));
    }
  };
  private onEditChange = (field: keyof IProduct, value: string): void => {
    this.setState((prev) => ({
      editItem: prev.editItem
        ? {
            ...prev.editItem,
            [field]: field === "ApprovalCost" ? Number(value) : value,
          }
        : null,
    }));
  };

  public render(): React.ReactElement<IHelloWorldProps> {
    const approvedCount = this.state.allEmplist.filter(
      (item) => item.Status === "Approved",
    ).length;

    const inProgressCount = this.state.allEmplist.filter(
      (item) => item.Status === "In Progress",
    ).length;

    const rejectedCount = this.state.allEmplist.filter(
      (item) => item.Status === "Rejected",
    ).length;
    const { emplist, currentPage, itemsPerPage } = this.state;

    const totalPages = Math.ceil(emplist.length / itemsPerPage);
    const indexOfLastItem = currentPage * itemsPerPage;
    const indexOfFirstItem = indexOfLastItem - itemsPerPage;

    const currentItems = emplist.slice(indexOfFirstItem, indexOfLastItem);

    return (
      <div className={styles.pageContainer}>
        <div className={styles.cardContainer}>
          {/* HEADER */}
          <div className={styles.headerRow}>
            <h2 className={styles.pageTitle}>Product Approval Management</h2>

            {/* <button className={styles.addButton} onClick={this.openAddModal}>
              <i className="bi bi-plus-lg"></i>
              Add Product
            </button> */}
          </div>
          {/* SUMMARY CARDS */}
          <div className="row mb-4">
            {/* Approved */}
            <div className="col-md-4">
              <div className={`${styles.cardBox} shadow-sm`}>
                <div className={`${styles.cardIcon} ${styles.approvedBG}`}>
                  <i className="bi bi-check-circle-fill"></i>
                </div>
                <div>
                  <h6 className="mb-1 text-muted">Approved</h6>
                  <h4 className="fw-bold">{approvedCount}</h4>
                </div>
              </div>
            </div>

            {/* In Progress */}
            <div className="col-md-4">
              <div className={`${styles.cardBox} shadow-sm`}>
                <div className={`${styles.cardIcon} ${styles.inProgressBG}`}>
                  <i className="bi bi-hourglass-split"></i>
                </div>
                <div>
                  <h6 className="mb-1 text-muted">In Progress</h6>
                  <h4 className="fw-bold">{inProgressCount}</h4>
                </div>
              </div>
            </div>

            {/* Rejected */}
            <div className="col-md-4">
              <div className={`${styles.cardBox} shadow-sm`}>
                <div className={`${styles.cardIcon} ${styles.rejectedBG}`}>
                  <i className="bi bi-x-circle-fill"></i>
                </div>
                <div>
                  <h6 className="mb-1 text-muted">Rejected</h6>
                  <h4 className="fw-bold">{rejectedCount}</h4>
                </div>
              </div>
            </div>
          </div>
          {/* SEARCH + ADD BUTTON ROW */}
          <div className="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
            {/* Search */}
            <div className={styles.searchContainer}>
              <i className={`bi bi-search ${styles.searchIcon}`} />

              <input
                className={styles.searchInput}
                placeholder="Search products..."
                onChange={this.onSearch}
              />
            </div>

            {/* Add Button */}
            <button className={styles.addButton} onClick={this.openAddModal}>
              <i className="bi bi-plus-lg"></i> Add Product
            </button>
          </div>
          {/* STATUS FILTER */}
          <div className="d-flex gap-2 mb-3 flex-wrap">
            {["All", "Approved", "In Progress", "Rejected", "Pending"].map(
              (status) => (
                <button
                  key={status}
                  className={`btn btn-sm ${
                    this.state.selectedStatus === status
                      ? "btn-primary"
                      : "btn-outline-primary"
                  }`}
                  onClick={() => this.onTabChange(status)}
                >
                  {status}
                </button>
              ),
            )}
          </div>
          {/* TABLE */}
          <table className={styles.table}>
            <thead className={styles.thead}>
              <tr>
                <th
                  onClick={() => this.onSort("ProductID")}
                  style={{ cursor: "pointer" }}
                  className={styles.th}
                >
                  Product ID {this.renderSortIcon("ProductID")}
                </th>
                <th
                  onClick={() => this.onSort("ProductName")}
                  style={{ cursor: "pointer" }}
                  className={styles.th}
                >
                  Product Name {this.renderSortIcon("ProductName")}
                </th>
                <th
                  onClick={() => this.onSort("Owner")}
                  style={{ cursor: "pointer" }}
                  className={styles.th}
                >
                  Owner {this.renderSortIcon("Owner")}
                </th>
                <th
                  onClick={() => this.onSort("ApprovalCost")}
                  style={{ cursor: "pointer" }}
                  className={styles.th}
                >
                  Approval Cost {this.renderSortIcon("ApprovalCost")}
                </th>
                <th
                  onClick={() => this.onSort("Status")}
                  style={{ cursor: "pointer" }}
                  className={styles.th}
                >
                  Status {this.renderSortIcon("Status")}
                </th>
                <th className={styles.th}></th>
                <th className={styles.th}></th>
              </tr>
            </thead>

            <tbody>
              {/* {emplist.map((item, index) => ( */}
              {currentItems.length > 0 ? (
                currentItems.map((item, index) => (
                  <tr key={item.Id} className={styles.tr}>
                    <td className={styles.td}>{item.ProductID}</td>
                    <td className={styles.td}>{item.ProductName}</td>
                    <td className={styles.td}>
                      <div className={styles.ownerContainer}>
                        <div
                          className={styles.avatar}
                          style={{
                            backgroundColor: `hsl(${(item.Owner?.length || 1) * 40}, 70%, 50%)`,
                          }}
                        >
                          {item.Owner
                            ? item.Owner.split(" ")
                                .map((n) => n[0])
                                .join("")
                                .toUpperCase()
                            : "NA"}
                        </div>
                        <span className={styles.ownerName}>{item.Owner}</span>
                      </div>
                    </td>{" "}
                    <td className={styles.td}>{item.ApprovalCost}</td>
                    <td className={styles.td}>
                      <span
                        className={`${styles.statusBadge} ${
                          item.Status === "Approved"
                            ? styles.approved
                            : item.Status === "In Progress"
                              ? styles.inProgress
                              : item.Status === "Rejected"
                                ? styles.rejected
                                : styles.pending
                        }`}
                      >
                        {item.Status}
                      </span>
                    </td>{" "}
                    <td className={styles.td}>
                      <i
                        className="bi bi-trash-fill text-danger"
                        style={{ cursor: "pointer", fontSize: "18px" }}
                        // onClick={() => this.openDeleteModal(item.ProductID)}
                        onClick={() => this.openDeleteModal(item.Id)}
                      ></i>
                    </td>
                    <td className={styles.td}>
                      <i
                        className="bi bi-pencil-square text-primary"
                        style={{ cursor: "pointer", fontSize: "18px" }}
                        onClick={() => this.openEditModal(item)}
                      ></i>
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td
                    className={styles.td}
                    colSpan={9}
                    style={{ textAlign: "center" }}
                  >
                    No records found.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
          {totalPages > 0 && (
            <div className="d-flex justify-content-between align-items-center mt-3 flex-wrap">
              {/* Left: record info */}
              <div className="text-muted small">
                Showing{" "}
                <strong>
                  {indexOfFirstItem + 1}–
                  {indexOfLastItem > emplist.length
                    ? emplist.length
                    : indexOfLastItem}
                </strong>{" "}
                of <strong>{emplist.length}</strong> records
              </div>

              {/* Right: pagination */}
              <nav>
                <ul
                  className={`pagination pagination-sm mb-0 ${styles.customPagination}`}
                >
                  {/* Previous */}
                  <li
                    className={`page-item ${currentPage === 1 ? "disabled" : ""}`}
                  >
                    <button
                      className="page-link"
                      onClick={this.prevPage}
                      aria-label="Previous"
                    >
                      ‹
                    </button>
                  </li>

                  {/* Page Numbers */}
                  {(() => {
                    const pages = [];
                    for (let i = 1; i <= totalPages; i++) {
                      pages.push(
                        <li
                          key={i}
                          className={`page-item ${
                            currentPage === i ? "active" : ""
                          }`}
                        >
                          <button
                            className="page-link"
                            onClick={() => this.goToPage(i)}
                          >
                            {i}
                          </button>
                        </li>,
                      );
                    }
                    return pages;
                  })()}

                  {/* Next */}
                  <li
                    className={`page-item ${
                      currentPage === totalPages ? "disabled" : ""
                    }`}
                  >
                    <button
                      className="page-link"
                      onClick={this.nextPage}
                      aria-label="Next"
                    >
                      ›
                    </button>
                  </li>
                </ul>
              </nav>
            </div>
          )}
        </div>

        {/* DELETE MODAL */}
        {this.state.showDeleteModal && (
          <>
            {/* Modal */}
            <div
              className="modal fade show d-block"
              tabIndex={-1}
              style={{ zIndex: 1055 }}
            >
              <div className="modal-dialog modal-dialog-centered">
                <div className="modal-content">
                  <div className="modal-header">
                    <h5 className="modal-title">Confirm Delete</h5>
                    <button
                      type="button"
                      className="btn-close"
                      onClick={this.closeDeleteModal}
                    ></button>
                  </div>

                  <div className="modal-body d-flex justify-content-center align-items-center">
                    <p>Are you sure you want to delete this record?</p>
                  </div>

                  <div className="modal-footer">
                    <button
                      className="btn btn-secondary"
                      onClick={this.closeDeleteModal}
                    >
                      Cancel
                    </button>
                    <button
                      className="btn btn-danger"
                      onClick={this.confirmDelete}
                    >
                      Delete
                    </button>
                  </div>
                </div>
              </div>
            </div>

            {/* Backdrop */}
            <div
              className="modal-backdrop fade show"
              style={{ zIndex: 1050 }}
            ></div>
          </>
        )}
        {/* EDIT / ADD MODAL */}
        {this.state.showEditModal && this.state.editItem && (
          <>
            {/* Modal */}
            <div
              className="modal fade show d-block"
              tabIndex={-1}
              style={{ zIndex: 1055 }}
            >
              <div className="modal-dialog modal-dialog-centered">
                <div className="modal-content">
                  {/* Header */}
                  <div className="modal-header">
                    <h5 className="modal-title">
                      {this.state.isEditMode ? "Edit Product" : "Add Product"}
                    </h5>

                    <button
                      className="btn-close"
                      onClick={this.closeEditModal}
                    />
                  </div>

                  {/* Body */}
                  <div className="modal-body">
                    <div className="row">
                      {/* Product ID */}
                      <div className="col-12 mb-3">
                        <label className="form-label fw-semibold">
                          Product ID
                        </label>

                        <input
                          className="form-control"
                          value={this.state.editItem.ProductID}
                          disabled={this.state.isEditMode}
                          onChange={(e) =>
                            this.onEditChange("ProductID", e.target.value)
                          }
                        />
                      </div>

                      {/* Product Name */}
                      <div className="col-12 mb-3">
                        <label className="form-label fw-semibold">
                          Product Name
                        </label>

                        <input
                          className="form-control"
                          value={this.state.editItem.ProductName}
                          onChange={(e) =>
                            this.onEditChange("ProductName", e.target.value)
                          }
                        />
                      </div>

                      {/* Product Manager */}
                      <div className="col-12 mb-3">
                        <label className="form-label fw-semibold">Owner</label>

                        <input
                          className="form-control"
                          value={this.state.editItem.Owner}
                          onChange={(e) =>
                            this.onEditChange("Owner", e.target.value)
                          }
                        />
                      </div>

                      {/* Approval Cost */}
                      <div className="col-12 mb-3">
                        <label className="form-label fw-semibold">
                          Approval Cost
                        </label>

                        <input
                          type="number"
                          className="form-control"
                          value={this.state.editItem.ApprovalCost}
                          onChange={(e) =>
                            this.onEditChange("ApprovalCost", e.target.value)
                          }
                        />
                      </div>

                      {/* Status */}
                      <div className="col-12 mb-3">
                        <label className="form-label fw-semibold">Status</label>

                        <select
                          className="form-control"
                          value={this.state.editItem.Status}
                          onChange={(e) =>
                            this.onEditChange("Status", e.target.value)
                          }
                        >
                          <option value="Pending">Pending</option>
                          <option value="In Progress">In Progress</option>
                          <option value="Approved">Approved</option>
                          <option value="Rejected">Rejected</option>
                        </select>
                      </div>
                    </div>
                  </div>

                  {/* Footer */}
                  <div className="modal-footer">
                    <button
                      className="btn btn-secondary"
                      onClick={this.closeEditModal}
                    >
                      Cancel
                    </button>

                    <button
                      className={`btn ${this.state.isEditMode ? "btn-primary" : "btn-success"}`}
                      onClick={() => {
                        if (this.state.isEditMode) {
                          this.updateItem();
                        } else {
                          this.addItem();
                        }
                      }}
                    >
                      {this.state.isEditMode ? "Update" : "Add"}
                    </button>
                  </div>
                </div>
              </div>
            </div>

            {/* Backdrop */}
            <div
              className="modal-backdrop fade show"
              style={{ zIndex: 1050 }}
            />
          </>
        )}
        {/* Toaster Message */}
        {this.state.showToast && (
          <div
            className="toast show position-fixed top-0 m-3"
            style={{ zIndex: 1000, minWidth: "250px" }}
          >
            <div
              className={`toast-body text-center fw-semibold text-white ${styles.toastBg}`}
            >
              {this.state.toastMessage}
            </div>
          </div>
        )}
      </div>
    );
  }
}
