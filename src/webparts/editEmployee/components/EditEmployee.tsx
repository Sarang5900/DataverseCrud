import * as React from 'react';
import { IEditEmployeeProps } from './IEditEmployeeProps';
import { TextField, DefaultButton, ComboBox, IComboBoxOption, Toggle, Stack, PrimaryButton, Label, IComboBox } from '@fluentui/react';
import ApiService from '../../../Services/ApiService';
import Swal from 'sweetalert2';

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

interface IEmployeeForm {
  employee: IEmployee;
  errors: Record<string, string>;
  teamMembers: IComboBoxOption[];
}

const params = new URLSearchParams(window.location.search);
// const employeeId = "450754b1-e5c2-ef11-b8e8-6045bd720e8b";
const employeeId = params.get('empId');

if (!employeeId || isNaN(Number(employeeId))) {
  console.error('Invalid employee ID in URL:', employeeId);
}

export default class EditEmployee extends React.Component<IEditEmployeeProps, IEmployeeForm> {
  private apiService: ApiService;

  constructor(props: IEditEmployeeProps) {
    super(props);
    this.state = {
      employee: {
        id: '',
        firstname: '',
        lastname: '',
        notes: '',
        title: '',
        age: 0,
        emailaddress: '',
        hiredate: '',
        goodattheirwork: false,
        selectedTeamMember: ''
      },
      errors: {},
      teamMembers: []
    };

    this.apiService = new ApiService(
      'a5281160-c870-4bde-b267-a52c2db0c107', 
      'https://login.microsoftonline.com/fbcab41b-0c15-41f2-9858-b64924a83a6c'
    );
  }

  async componentDidMount() {
    try {
      if (employeeId) {
        const teamMembers = await this.apiService.getTeamMembers(this.props.context);
        const teamMemberOptions: IComboBoxOption[] = teamMembers.map((member) => ({
          key: member.id,  // team member's ID (GUID)
          text: member.name,  // team member's name
        }));
  
        this.setState({ teamMembers: teamMemberOptions });
  
        await this.fetchEmployeeDetails(employeeId);
      }
    } catch (error) {
      console.error('Error initializing component:', error);
    }
  }

  private async fetchEmployeeDetails(employeeId: string): Promise<void> {
    try {
      const employee = await this.apiService.fetchEmployeeById(this.props.context, employeeId);
  
      console.log(employee);
      
      if (employee) {
        if (employee.hiredate) {
          const hiredate = new Date(employee.hiredate).toISOString().split('T')[0];
          employee.hiredate = hiredate;
        }
  
        const titleOptions = [
          { key: 100000000, text: "Developer" },
          { key: 100000001, text: "Manager" },
          { key: 100000002, text: "Designer" },
          { key: 100000003, text: "Analyst" },
        ];
  
        const matchedOption = titleOptions.find(option => option.text === employee.title);
        employee.title = matchedOption ? (matchedOption.key as unknown as string) : '';

        const matchedTeamMember = this.state.teamMembers.find(
          (member) => member.text === employee.selectedTeamMember
        );
        console.log(matchedTeamMember);
        
        employee.selectedTeamMember = matchedTeamMember?.key as string;
  
        this.setState({
          employee: employee,
        });
      } else {
        console.error(`Employee with ID ${employeeId} not found`);
      }
    } catch (error) {
      console.error('Error fetching employee details:', error);
      alert('Failed to fetch employee details');
    }
  }

    private handleTeamMemberChange = (event: React.FormEvent<IComboBox>, option?: IComboBoxOption) => {
      if (option) {
        this.setState((prevState) => ({
          employee: {
            ...prevState.employee, 
            selectedTeamMember: option.key as string, 
          },
        }));
      }
    };

  private _handleSubmit = async (event: React.MouseEvent<HTMLButtonElement>): Promise<void> => {
    event.preventDefault();

    const errors: Record<string, string> = {};
    const { firstname, lastname, notes, title, age, emailaddress, hiredate, selectedTeamMember } = this.state.employee;

    if (!firstname) errors.firstname = 'First name is required';
    if (!lastname) errors.lastname = 'Last name is required';
    if (!notes) errors.notes = 'Notes are required';
    if (!hiredate) errors.hiredate = 'Hire date is required';
    if (!title) errors.title = 'Title is required';
    if (!age) errors.age = 'Age is required';
    if (!emailaddress) errors.emailaddress = 'Email address is required';
    if (!selectedTeamMember) errors.selectedTeamMember = 'Team member is required';

    if (Object.keys(errors).length > 0) {
      this.setState({ errors });
      return; 
    }
  
    try {
      const result = await Swal.fire({
        title: 'Are you sure?',
        text: 'Do you want to update the employee details?',
        icon: 'warning',
        showCancelButton: true,
        confirmButtonText: 'Yes, update it!',
        cancelButtonText: 'No, cancel',
      });
  
      if (!result.isConfirmed) {
        return; 
      }
  
      const updatedEmployee: IEmployee = { ...this.state.employee };
  
      const context = this.props.context;
      await this.apiService.editEmployee(context, updatedEmployee);
  
      await Swal.fire({
        title: 'Updated!',
        text: 'Employee details updated successfully.',
        icon: 'success',
        confirmButtonText: 'Ok',
      });

      this.handleGoToList();
    } catch (error) {
      console.error('Error updating employee:', error);
  
      await Swal.fire({
        title: 'Error!',
        text: 'An error occurred while updating the employee. Please try again.',
        icon: 'error',
        confirmButtonText: 'Ok',
      });
    }
  };
  
  private handleInputChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>, field: keyof IEmployee): void => {
    const { value } = e.target;
    this.setState((prevState) => ({
      employee: {
        ...prevState.employee,
        [field]: value,
      },
    }));
  };

  private handleToggleChange = (checked: boolean): void => {
    this.setState({
      employee: {
        ...this.state.employee,
        goodattheirwork: checked,
      },
    });
  };

  private handleTitleChange = (
    event: React.FormEvent<IComboBox>, 
    option?: IComboBoxOption, 
    index?: number, 
    value?: string
  ): void => {
    this.setState({
      employee: {
        ...this.state.employee,
        title: option?.key as string || '',
      },
    });
  };

  handleGoToList = (): void => {
    const gotoList = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/EmployeeDataList.aspx`;
    window.location.href = gotoList;
  };

  render() {
    const titleOptions: IComboBoxOption[] = [
      { key: 100000000, text: "Developer" },
      { key: 100000001, text: "Manager" },
      { key: 100000002, text: "Designer" },
      { key: 100000003, text: "Analyst" },
    ];

    const { employee, errors, teamMembers } = this.state;
    const { firstname, lastname, notes, title, age, emailaddress, hiredate, goodattheirwork,  } = employee;

    return (
      <div
        style={{
          backgroundPosition: 'center',
          backgroundSize: 'cover',
          height: '100%',
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
          flexDirection: 'column',
        }}
      >
        <Label
          styles={{
            root: {
              marginTop: '20px',
              marginBottom: '20px',
              textAlign: 'center',
            },
          }}
        >
          <h1>Edit Employee</h1>
        </Label>
        <Stack
          tokens={{ childrenGap: 20 }}
          styles={{
            root: {
              width: '100%',
              maxHeight: '100vh',
              maxWidth: 700,
              margin: '20px 20px',
              padding: 40,
              borderRadius: '12px',
              backdropFilter: 'blur(10px) saturate(180%)',
              WebkitBackdropFilter: 'blur(16px) saturate(180%)',
              border: '1px solid rgba(255, 255, 255, 0.125)',
              boxShadow: '12px 12px 12px 12px rgba(0, 0, 0, 0.2)',
              display: 'flex',
              flexDirection: 'column',
              overflowY: 'auto',
              flexShrink: 0,
            },
          }}
        >
          <Stack horizontal tokens={{ childrenGap: 15 }}>
            <Stack.Item styles={{ root: { width: '50%' } }}>
              <TextField
                label="First Name"
                name="firstname"
                value={firstname}
                onChange={(e) => this.handleInputChange(e as React.ChangeEvent<HTMLInputElement>, 'firstname')}
                errorMessage={errors?.firstname}
                required
                styles={{
                  fieldGroup: {
                    backgroundColor: 'transparent',
                  },
                }}
              />
            </Stack.Item>
            <Stack.Item styles={{ root: { width: '50%' } }}>
              <TextField
                label="Last Name"
                name="lastname"
                value={lastname}
                onChange={(e) => this.handleInputChange(e as React.ChangeEvent<HTMLInputElement>, 'lastname')}
                errorMessage={errors?.lastname}
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
            name="notes"
            multiline
            value={notes}
            errorMessage={errors?.notes}
            onChange={(e) => this.handleInputChange(e as React.ChangeEvent<HTMLInputElement>, 'notes')}
            required
            styles={{
              fieldGroup: {
                backgroundColor: 'transparent',
              },
            }}
          />

          <Stack horizontal tokens={{ childrenGap: 15 }}>
            <Stack.Item styles={{ root: { width: '50%' } }}>
              <TextField
                label="Hire Date"
                name="hiredate"
                type="date"
                value={hiredate}
                onChange={(e) => this.handleInputChange(e as React.ChangeEvent<HTMLInputElement>, 'hiredate')}
                errorMessage={errors?.hiredate}
                required
                styles={{
                  fieldGroup: {
                    backgroundColor: 'transparent',
                  },
                }}
              />
            </Stack.Item>
            <Stack.Item styles={{ root: { width: '50%' } }}>
              <ComboBox
                label="Title"
                options={titleOptions}
                selectedKey={title}
                onChange={this.handleTitleChange}
                errorMessage={errors?.title}
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
            <Stack.Item styles={{ root: { width: '50%' } }}>
              <TextField
                label="Age"
                name="age"
                type="number"
                value={age.toString()}
                onChange={(e) => this.handleInputChange(e as React.ChangeEvent<HTMLInputElement>, 'age')}
                errorMessage={errors?.age}
                required
                styles={{
                  fieldGroup: {
                    backgroundColor: 'transparent',
                  },
                }}
              />
            </Stack.Item>
            <Stack.Item styles={{ root: { width: '50%' } }}>
              <TextField
                label="Email Address"
                name="emailaddress"
                value={emailaddress}
                onChange={(e) => this.handleInputChange(e as React.ChangeEvent<HTMLInputElement>, 'emailaddress')}
                errorMessage={errors?.emailaddress}
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
            <Stack.Item styles={{ root: { width: '50%' } }}>
              <ComboBox
                label="Select Team Member"
                options={teamMembers}
                selectedKey={employee.selectedTeamMember}
                onChange={this.handleTeamMemberChange}
                required
                allowFreeform
                autoComplete="on"
                placeholder="Search for a team member"
                errorMessage={errors?.selectedTeamMember}
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

            <Stack.Item styles={{ root: { width: '50%' } }}>
              <Toggle
                label="Good at their work"
                checked={goodattheirwork}
                onChange={(_, checked) => this.handleToggleChange(checked || false)}
              />
            </Stack.Item>
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 10 }}>
            <PrimaryButton text="Save Employee" onClick={this._handleSubmit} />
            <DefaultButton
              text="Cancel"
              onClick={() =>
                this.setState({
                  employee: {
                    id: '',
                    firstname: '',
                    lastname: '',
                    notes: '',
                    title: '',
                    age: 0,
                    emailaddress: '',
                    hiredate: '',
                    goodattheirwork: false,
                    selectedTeamMember: '',
                  },
                }, this.handleGoToList)
              }
            />
          </Stack>
        </Stack>
      </div>
    );
  }
}
