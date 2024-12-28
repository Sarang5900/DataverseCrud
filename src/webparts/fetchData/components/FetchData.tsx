import * as React from "react";
import { IFetchDataProps } from "./IFetchDataProps";
import {
  DetailsList,
  PrimaryButton,
  DefaultButton,
  Stack,
  SearchBox,
  IColumn,
  MessageBar,
  MessageBarType,
} from "@fluentui/react";
import Swal from 'sweetalert2';
import styles from "./FetchData.module.scss";
import ApiService from "../../../Services/ApiService";

interface IEmployee {
  id: string;
  firstname: string;
  lastname: string;
  notes: string;
  title: string;
  age: number;
  emailaddress: string;
  hiredate: string;
  goodattheirwork: boolean;
  selectedTeamMember: string;
}

interface IFetchDataState {
  employees: IEmployee[];
  loading: boolean;
  error: string | undefined;
  columns: IColumn[];
  sortedColumnKey: keyof IEmployee | undefined;
  isSortedDescending: boolean;
  currentPage: number;
  pageSize: number;
  noSearchResults: boolean;
}


export default class FetchData extends React.Component<IFetchDataProps,IFetchDataState> {
   private originalEmployees: IEmployee[] = [];
  private employeeService: ApiService;

  constructor(props: IFetchDataProps) {
    super(props);
    this.state = {
      employees: [],
      loading: true,
      error: undefined,
      columns: [],
      sortedColumnKey: undefined,
      isSortedDescending: false,
      currentPage: 1,
      pageSize: 5,
      noSearchResults: false,
    };
    this.employeeService = new ApiService(
      'a5281160-c870-4bde-b267-a52c2db0c107', 
      'https://login.microsoftonline.com/fbcab41b-0c15-41f2-9858-b64924a83a6c'
    );
  }


  public async componentDidMount(): Promise<void> {
    try {
      const employees = await this.employeeService.getEmployees(this.props.context);
      
      this.originalEmployees = employees;
      this.setState({ employees, loading: false, columns: this._getColumns() });
    } catch (error) {
      this.setState({
        loading: false,
        error: `Error fetching employees: ${error.message}`,
      });
    }
  }
  

  handleEditEmployee = (empId: string): void => {
    const editTaskUrl = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/EditEmployee.aspx?empId=${empId}`;
    window.location.href = editTaskUrl;
  };

  handleAddEmployee = (): void => {
    const editTaskUrl = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/AddEmployee.aspx`;
    window.location.href = editTaskUrl;
  };


  private async deleteEmployee(employeeId: string): Promise<void> {
    try {
      const result = await Swal.fire({
        title: "Are you sure?",
        text: "You won't be able to revert this!",
        icon: "warning",
        showCancelButton: true,
        confirmButtonColor: "#3085d6",
        cancelButtonColor: "#d33",
        confirmButtonText: "Yes, delete it!",
      });
  
      if (result.isConfirmed) {
        await this.employeeService.deleteEmployee(this.props.context, employeeId);
        
        this.setState({
          employees: this.state.employees.filter((employee) => employee.id !== employeeId),
        });

        // const employees = await this.employeeService.getEmployees(this.props.context);
      
        // this.originalEmployees = employees;
        // this.setState({ employees, loading: false, columns: this._getColumns() });
  
        await Swal.fire({
          title: "Deleted!",
          text: "The employee record has been deleted.",
          icon: "success",
          confirmButtonColor: "#3085d6",
        });
      }
    } catch (error) {
      await Swal.fire({
        title: "Error",
        text: `Failed to delete employee: ${error.message}`,
        icon: "error",
        confirmButtonColor: "#3085d6",
      });
    }
  }
  


  private searchEmployees = (searchText: string): void => {
    if (!searchText) {
      this.setState({
        employees: this.originalEmployees,
        noSearchResults: false,
        currentPage: 1,
      });
      return;
    }

    const filteredEmployees = this.originalEmployees.filter((employee) =>
      Object.values(employee).some((value) =>
        value?.toString().toLowerCase().includes(searchText.toLowerCase())
      )
    );

    this.setState({
      employees: filteredEmployees,
      noSearchResults: filteredEmployees.length === 0,
      currentPage: 1,
    });
  };

  private handlePageChange = (pageNumber: number): void => {
    this.setState({ currentPage: pageNumber });
  };

  private _getColumns = (): IColumn[] => [
    {
      key: "firstname",
      name: "First Name",
      fieldName: "firstname",
      minWidth: 100,
      isResizable: true,
      isSorted: this.state.sortedColumnKey === 'firstname',
      isSortedDescending: this.state.isSortedDescending,
      onColumnClick: (event, column) => this.onColumnClick(column),
    },
    {
      key: "lastname",
      name: "Last Name",
      fieldName: "lastname",
      minWidth: 100,
      isResizable: true,
      isSorted: this.state.sortedColumnKey === 'lastname',
      isSortedDescending: this.state.isSortedDescending,
      onColumnClick: (event, column) => this.onColumnClick(column),
    },
    {
      key: "title",
      name: "Title",
      fieldName: "title",
      minWidth: 100,
      isResizable: true,
      isSorted: this.state.sortedColumnKey === 'title',
      isSortedDescending: this.state.isSortedDescending,
      onColumnClick: (event, column) => this.onColumnClick(column),
    },
    {
      key: "age",
      name: "Age",
      fieldName: "age",
      minWidth: 50,
      isResizable: true,
      isSorted: this.state.sortedColumnKey === 'age',
      isSortedDescending: this.state.isSortedDescending,
      onColumnClick: (event, column) => this.onColumnClick(column),
    },
    {
      key: "emailaddress",
      name: "Email Address",
      fieldName: "emailaddress",
      minWidth: 150,
      isResizable: true,
      isSorted: this.state.sortedColumnKey === 'emailaddress',
      isSortedDescending: this.state.isSortedDescending,
      onColumnClick: (event, column) => this.onColumnClick(column),
    },
    {
      key: "hiredate",
      name: "Hire Date",
      fieldName: "hiredate",
      minWidth: 100,
      isResizable: true,
      isSorted: this.state.sortedColumnKey === 'hiredate',
      isSortedDescending: this.state.isSortedDescending,
      onColumnClick: (event, column) => this.onColumnClick(column),
    },
    {
      key: "goodattheirwork",
      name: "Good At Work",
      fieldName: "goodattheirwork",
      minWidth: 50,
      isResizable: true,
      onRender: (item: IEmployee) => (item.goodattheirwork ? "Yes" : "No"),
      isSorted: this.state.sortedColumnKey === 'goodattheirwork',
      isSortedDescending: this.state.isSortedDescending,
      onColumnClick: (event, column) => this.onColumnClick(column),
    },
    {
      key: "selectedTeamMember",
      name: "Team Member",
      fieldName: "selectedTeamMember",
      minWidth: 100,
      isResizable: true,
      isSorted: this.state.sortedColumnKey === 'selectedTeamMember',
      isSortedDescending: this.state.isSortedDescending,
      onColumnClick: (event, column) => this.onColumnClick(column),
    },
    {
      key: "notes",
      name: "Notes",
      fieldName: "notes",
      minWidth: 150,
      isResizable: true,
      isSorted: this.state.sortedColumnKey === 'notes',
      isSortedDescending: this.state.isSortedDescending,
      onColumnClick: (event, column) => this.onColumnClick(column),
    },
    {
      key: "actions",
      name: "Actions",
      fieldName: "actions",
      minWidth: 200,
      onRender: (item: IEmployee) => (
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton text="Edit" onClick={() => this.handleEditEmployee(item.id)} />
          <DefaultButton
            text="Delete"
            onClick={() => this.deleteEmployee(item.id)}
            styles={{ root: { backgroundColor: "#d32f2f", color: "#fff" } }}
          />
        </Stack>
      ),
    },
  ];
  

  private onColumnClick(column: IColumn): void {
    const { sortedColumnKey, isSortedDescending } = this.state;
    const newSortDirection = sortedColumnKey === column.fieldName && isSortedDescending ? 'ascending' : 'descending';
    const sortedEmployees = this.sortEmployees(this.state.employees, column.fieldName as keyof IEmployee, newSortDirection === 'descending');
    this.setState({
      sortedColumnKey: column.fieldName as keyof IEmployee,
      isSortedDescending: newSortDirection === 'descending',
      employees: sortedEmployees,
    });
  }

  private sortEmployees(employees: IEmployee[], columnKey: keyof IEmployee, descending: boolean): IEmployee[] {
    const sortedEmployees = [...employees].sort((a, b) => {
      const valueA = a[columnKey];
      const valueB = b[columnKey];
  
      const stringA = typeof valueA === "string" ? valueA.toLowerCase() : valueA;
      const stringB = typeof valueB === "string" ? valueB.toLowerCase() : valueB;
  
      // Compare values
      if (stringA < stringB) return descending ? 1 : -1;
      if (stringA > stringB) return descending ? -1 : 1;
      return 0;
    });
    return sortedEmployees;
  }
  

  public render(): React.ReactNode {
    const { loading, employees, error, currentPage, pageSize, noSearchResults } = this.state;

    if (loading) {
      return <div>Loading...</div>;
    }

    if (error) {
      return <div>Error: {error}</div>;
    }

    const startIndex = (currentPage - 1) * pageSize;
    const paginatedEmployees = employees.slice(startIndex, startIndex + pageSize);
    const totalPages = Math.ceil(employees.length / pageSize);

    return (
      <div>
        <h1>Employee Data</h1>

        <header className={styles.header}>
          <SearchBox
            placeholder="Search here..."
            onChange={(_, newValue) => this.searchEmployees(newValue || "")}
            styles={{
              root: {
                width: "100%",
                maxWidth: "250px",
                backgroundColor: 'rgba(255, 255, 255, 0.8)',
                borderRadius: '5px',
                alignItems: 'center',
              },
            }}
          />

          <button className={styles.addButton} onClick={this.handleAddEmployee}>
            <svg
              height="24"
              width="24"
              viewBox="0 0 24 24"
              xmlns="http://www.w3.org/2000/svg"
            >
              <path d="M0 0h24v24H0z" fill="none"/>
              <path d="M11 11V5h2v6h6v2h-6v6h-2v-6H5v-2z" fill="currentColor"/>
            </svg>
            <span>Add Employee</span>
          </button>
        </header>

        <DetailsList
          items={paginatedEmployees}
          columns={this._getColumns()}
          setKey="set"
          layoutMode={1}
          checkboxVisibility={2}
        />

        {noSearchResults && (
          <MessageBar messageBarType={MessageBarType.warning}>
            No matched record found.
          </MessageBar>
        )}

        <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginTop: "20px" }}>

          <PrimaryButton
            text="Previous"
            onClick={() => this.handlePageChange(currentPage - 1)}
            disabled={currentPage === 1} 
            style={{
              backgroundColor: currentPage === 1 ? "#f3f2f1" : undefined,
              color: currentPage === 1 ? "#a19f9d" : undefined,
            }}
          />

          <span style={{ alignSelf: "center", fontWeight: "bold" }}>
            Page {currentPage} of {totalPages}
          </span>

          <PrimaryButton
            text="Next"
            onClick={() => this.handlePageChange(currentPage + 1)}
            disabled={currentPage === totalPages} 
            style={{
              backgroundColor: currentPage === totalPages ? "#f3f2f1" : undefined,
              color: currentPage === totalPages ? "#a19f9d" : undefined,
            }}
          />
        </Stack>

      </div>
    );
  }
}
