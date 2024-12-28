import * as React from "react";
import {
  TextField,
  PrimaryButton,
  DefaultButton,
  Stack,
  Toggle,
  Label,
  IComboBoxOption,
  ComboBox,
  IComboBox,
} from "@fluentui/react";
import type { IAddEmployeeProps } from "./IAddEmployeeProps";
import ApiService from "../../../Services/ApiService";
import styles from "./AddEmployee.module.scss";
import Swal from "sweetalert2";


interface IEmployee {
  id: string;
  firstname: string;
  lastname: string;
  notes: string
  title: string;
  age: number;
  emailaddress: string;
  hiredate: string;
  goodattheirwork: boolean;
  selectedTeamMember: string;
}

interface IAddEmployeeState {
  employee: Omit<IEmployee, "id">; // All fields except `id`
  errors: { [key in keyof IEmployee]?: string };
  success: string | undefined;
  teamMembers: IComboBoxOption[];
}

const titleOptions: IComboBoxOption[] = [
  { key: 100000000, text: "Developer" },
  { key: 100000001, text: "Manager" },
  { key: 100000002, text: "Designer" },
  { key: 100000003, text: "Analyst" },
];

export default class AddEmployee extends React.Component<
  IAddEmployeeProps,
  IAddEmployeeState
> {
  private employeeService: ApiService;

  constructor(props: IAddEmployeeProps) {
    super(props);

    this.state = {
      employee: {
        firstname: "",
        lastname: "",
        notes: "",
        title: "",
        age: 0,
        emailaddress: "",
        hiredate: "",
        goodattheirwork: false,
        selectedTeamMember: '',
      },
      errors: {},
      success: undefined,       
      teamMembers: [],
    };

    this.employeeService = new ApiService(
      "a5281160-c870-4bde-b267-a52c2db0c107",
      "https://login.microsoftonline.com/fbcab41b-0c15-41f2-9858-b64924a83a6c"
    );
  }

  async componentDidMount() {
    try {
      const teamMembers = await this.employeeService.getTeamMembers(this.props.context);
      const teamMemberOptions: IComboBoxOption[] = teamMembers.map((member) => ({
        key: member.id,  // team member's ID (GUID)
        text: member.name,  // team member's name
      }));
      this.setState({ teamMembers: teamMemberOptions });
    } catch (error) {
      console.error('Error fetching team members:', error);
    }
  }
  
  private handleInputChange = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    field: keyof IEmployee
  ) => {
    const value = event.currentTarget.value;
    this.setState((prevState) => ({
      employee: {
        ...prevState.employee,
        [field]: value,
      },
      errors: {
        ...prevState.errors,
        [field]: undefined,
      },
    }));
  };

  private handleToggleChange = (checked: boolean): void => {
    this.setState((prevState) => ({
      employee: { ...prevState.employee, goodattheirwork: checked },
      errors: { ...prevState.errors, goodattheirwork: undefined },
    }));
  };

  private handleTitleChange = (
    event: React.FormEvent<IComboBox>,
    option?: IComboBoxOption
  ) => {
    if (option) {
      this.setState((prevState) => ({
        employee: {
          ...prevState.employee,
          title: option.key as string,
        },
        errors: { ...prevState.errors, title: undefined }, 
      }));
    }
  };

  private handleTeamMemberChange = (event: React.FormEvent<IComboBox>, option?: IComboBoxOption) => {
    if (option) {
      this.setState((prevState) => ({
        employee: {
          ...prevState.employee, 
          selectedTeamMember: option.key as string, 
        },
        errors: { ...prevState.errors, selectedTeamMember: undefined },
      }));
    }
  };

  private handleSubmit = async (): Promise<void> => {
    const { employee } = this.state;
    const errors: { [key in keyof IEmployee]?: string } = {};

    if (!employee.firstname) errors.firstname = "First name is required.";
    if (!employee.lastname) errors.lastname = "Last name is required.";
    if (!employee.title) errors.title = "Title is required.";
    if (!employee.notes) errors.notes = "Notes is required.";
    if (!employee.age) errors.age = "Age is required.";
    if (!employee.emailaddress) errors.emailaddress = "Email address is required.";
    if (!employee.hiredate) errors.hiredate = "Hire date is required.";

    if (employee.hiredate && new Date(employee.hiredate) > new Date()) {
        errors.hiredate = "Hire date cannot be in the future.";
    }

    if (employee.age && (employee.age < 0 || employee.age < 18)) {
        errors.age = "Age should be at least 18 and cannot be less than 0.";
    }else if(employee.age.toString() !== "" && isNaN(employee.age)){
      errors.age = "Age should be a number.";
    }

    const emailRegex = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;
    if (employee.emailaddress && !emailRegex.test(employee.emailaddress)) {
        errors.emailaddress = "Please enter a valid email address.";
    }

    if(!employee.selectedTeamMember) errors.selectedTeamMember = "Please select a team member.";

    if (Object.keys(errors).length > 0) {
        this.setState({ errors, success: undefined });
        return;
    }
    const permissionGranted = await Swal.fire({
        title: 'Are you sure?',
        text: 'Do you want to add this employee?',
        icon: 'question',
        showCancelButton: true,
        confirmButtonText: 'Yes, add employee',
        cancelButtonText: 'Cancel'
    });

    if (!permissionGranted.isConfirmed) {
        return;
    }

    try {
        await this.employeeService.addEmployee(employee);

        await Swal.fire({
            icon: 'success',
            title: 'Success!',
            text: 'Employee added successfully!',
            confirmButtonText: 'OK'
        });

        this.setState({
            employee: {
                firstname: "",
                lastname: "",
                notes: "",
                title: "",
                age: 0,
                emailaddress: "",
                hiredate: "",
                goodattheirwork: false,
                selectedTeamMember: "",
            },
            errors: {},
            success: "Employee added successfully!",
        });
    } catch (error) {
        await Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: `Failed to add employee: ${error.message}`,
            confirmButtonText: 'Try Again'
        });

        this.setState({ errors: {}, success: undefined });
        console.error("Failed to add employee:", error);
    }
  };

  handleGoToList = (): void => {
    const gotoList = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/EmployeeDataList.aspx`;
    window.location.href = gotoList;
  };
  

  public render(): React.ReactElement<IAddEmployeeProps> {
    const { employee, errors, teamMembers,  } = this.state;

    return (
      <div
        style={{
          backgroundPosition: "center",
          backgroundSize: "cover",
          height: "100%",
          display: "flex",
          justifyContent: "center",
          alignItems: "center",
          flexDirection: "column",
        }}
      >
        <div className={styles.container}>
          <button className={styles.goToList} onClick={this.handleGoToList}>
            <span style={{ paddingTop: "5px", }}>Go To List</span>
            <svg
              xmlns="http://www.w3.org/2000/svg"
              fill="none"
              viewBox="0 0 74 74"
              height="34"
              width="34"
            >
              <circle strokeWidth="3" stroke="black" r="35.5" cy="37" cx="37" />
              <path
                fill="black"
                d="M25 35.5C24.1716 35.5 23.5 36.1716 23.5 37C23.5 37.8284 24.1716 38.5 25 38.5V35.5ZM49.0607 38.0607C49.6464 37.4749 49.6464 36.5251 49.0607 35.9393L39.5147 26.3934C38.9289 25.8076 37.9792 25.8076 37.3934 26.3934C36.8076 26.9792 36.8076 27.9289 37.3934 28.5147L45.8787 37L37.3934 45.4853C36.8076 46.0711 36.8076 47.0208 37.3934 47.6066C37.9792 48.1924 38.9289 48.1924 39.5147 47.6066L49.0607 38.0607ZM25 38.5L48 38.5V35.5L25 35.5V38.5Z"
              />
            </svg>
          </button>
        </div>
        <Label
          styles={{
            root: {
              marginTop: "20px",
              marginBottom: "20px",
              textAlign: "center",
            },
          }}
        >
          <h1>Add Employee</h1>
        </Label>
        <Stack
          tokens={{ childrenGap: 20 }}
          styles={{
            root: {
              width: "100%",
              maxHeight: "100vh",
              maxWidth: 700,
              margin: "20px 20px",
              padding: 40,
              borderRadius: "12px",
              backdropFilter: "blur(10px) saturate(180%)",
              WebkitBackdropFilter: "blur(16px) saturate(180%)",
              border: "1px solid rgba(255, 255, 255, 0.125)",
              boxShadow: "12px 12px 12px 12px rgba(0, 0, 0, 0.2)",
              display: "flex",
              flexDirection: "column",
              overflowY: "auto",
              flexShrink: 0,
            },
          }}  
        >
          <Stack horizontal tokens={{ childrenGap: 15 }}>
            <Stack.Item styles={{ root: { width: "50%" } }}>
              <TextField
                placeholder="Enter your first name"
                label="First Name"
                name="firstname"
                value={employee.firstname}
                onChange={(e) => this.handleInputChange(e, "firstname")}
                errorMessage={errors.firstname}
                required
                styles={{
                  fieldGroup: {
                    backgroundColor: 'transparent',
                  },
                }}
              />
            </Stack.Item>
            <Stack.Item styles={{ root: { width: "50%" } }}>
              <TextField
                placeholder="Enter your last name"
                label="Last Name"
                name="lastname"
                value={employee.lastname}
                onChange={(e) => this.handleInputChange(e, "lastname")}
                errorMessage={errors.lastname}
                required
                styles={{
                  fieldGroup: {
                    backgroundColor: 'transparent',
                  },
                }}
              />
            </Stack.Item>
          </Stack>

          <TextField
            label="Notes"
            placeholder="Enter brief description about notes"
            name="notes"
            multiline
            required
            value={employee.notes}
            errorMessage={errors.notes}
            onChange={(e) => this.handleInputChange(e, "notes")}
            styles={{
              fieldGroup: {
                backgroundColor: 'transparent',
              },
            }}
          />

          <Stack horizontal tokens={{ childrenGap: 15 }}>
            <Stack.Item styles={{ root: { width: "50%" } }}>
              <TextField
                label="Hire Date"
                name="hiredate"
                type="date"
                value={employee.hiredate}
                onChange={(e) => this.handleInputChange(e, "hiredate")}
                errorMessage={errors.hiredate}
                required
                styles={{
                  fieldGroup: {
                    backgroundColor: 'transparent',
                  },
                }}
              />
            </Stack.Item>
            <Stack.Item styles={{ root: { width: "50%" } }}>
              <ComboBox
                placeholder="Select title"
                label="Title"
                options={titleOptions}
                selectedKey={employee.title}
                onChange={this.handleTitleChange}
                errorMessage={errors.title}
                required
                styles={{
                  callout: {
                    minWidth: 300,
                    maxWidth: 500,
                    fontSize: '16px',
                    backgroundColor: 'transparent',  
                  },
                  root: {
                    width: '100%',
                    backgroundColor: 'transparent', 
                  },
                  input: {
                    backgroundColor: 'transparent', 
                  },
                }}
              />
            </Stack.Item>
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 15 }}>
            <Stack.Item styles={{ root: { width: "50%" } }}>
              <TextField
                placeholder="Enter your age"
                label="Age"
                name="age"
                value={employee.age === 0 ? "": employee.age.toString()}
                onChange={(e) => this.handleInputChange(e, "age")}
                errorMessage={errors.age}
                required
                styles={{
                  fieldGroup: {
                    backgroundColor: 'transparent',
                  },
                }}
              />
            </Stack.Item>
            <Stack.Item styles={{ root: { width: "50%" } }}>
              <TextField
                placeholder="Enter your email address"
                label="Email Address"
                name="emailaddress"
                value={employee.emailaddress}
                onChange={(e) => this.handleInputChange(e, "emailaddress")}
                errorMessage={errors.emailaddress}
                required
                styles={{
                  fieldGroup: {
                    backgroundColor: 'transparent',
                  },
                }}
              />
            </Stack.Item>
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 15 }}>
            <Stack.Item styles={{ root: { width: "50%" } }}>
            <ComboBox
              label="Select Team Member"
              options={teamMembers}
              selectedKey={employee.selectedTeamMember}
              onChange={this.handleTeamMemberChange}
              required
              allowFreeform
              placeholder="Search for a team member"
              errorMessage={errors.selectedTeamMember}
              styles={{
                callout: {
                  minWidth: 300,
                  maxWidth: 500,
                  fontSize: '16px',
                  backgroundColor: 'transparent',  
                },
                root: {
                  width: '100%',
                  backgroundColor: 'transparent', 
                },
                input: {
                  backgroundColor: 'transparent', 
                },
              }}
            />
            </Stack.Item>

            <Stack.Item styles={{ root: { width: "50%" } }}>
              <Toggle
                label="Good at their work"
                checked={employee.goodattheirwork}
                onChange={(_, checked) =>
                  this.handleToggleChange(checked || false)
                }
                styles={{
                  root: {
                    backgroundColor: 'transparent',
                  },
                }}
              />
            </Stack.Item>
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 10 }}>
            <PrimaryButton text="Add Employee" onClick={this.handleSubmit} />
            <DefaultButton
              text="Reset"
              onClick={() =>
                this.setState({
                  employee: {
                    firstname: "",
                    lastname: "",
                    notes: "",
                    title: "",
                    age: 0,
                    emailaddress: "",
                    hiredate: "",
                    goodattheirwork: false,
                    selectedTeamMember: "",
                  },
                  errors: {},
                  success: undefined,
                })
              }
            />
          </Stack>
        </Stack>
      </div>
    );
  }
}
