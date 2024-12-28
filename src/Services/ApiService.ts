import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { PublicClientApplication, AuthenticationResult, BrowserAuthError } from '@azure/msal-browser';

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

class ApiService {
  private msalInstance: PublicClientApplication;
  private apiUrl: string = "https://orgf23feab4.crm8.dynamics.com/api/data/v9.2";
  
  constructor(clientId: string, authority: string) {
    this.msalInstance = new PublicClientApplication({
      auth: {
        clientId: clientId,
        authority: authority,
        redirectUri: window.location.origin,
      },
    });
  }

  private async signIn(): Promise<void> {
    try {
      const loginResponse: AuthenticationResult = await this.msalInstance.loginPopup({
        scopes: ['https://orgf23feab4.crm8.dynamics.com/.default'], 
      });
      this.msalInstance.setActiveAccount(loginResponse.account);
    } catch (error) {
      console.error('Error during sign-in:', error);
    }
  }

  private async getAccessToken(): Promise<string> {
    await this.msalInstance.initialize();
    const request = { scopes: ['https://orgf23feab4.crm8.dynamics.com/.default'] };

    try {
      const activeAccount = this.msalInstance.getAllAccounts()[0];
      if (!activeAccount) {
        await this.signIn();
      } else {
        this.msalInstance.setActiveAccount(activeAccount);
      }

      const response: AuthenticationResult = await this.msalInstance.acquireTokenSilent(request);
      return response.accessToken;
    } catch (error) {
      console.error('Error acquiring token silently:', error);
      if (error instanceof BrowserAuthError && error.errorCode === "interaction_required") {
        const interactiveResponse = await this.msalInstance.acquireTokenPopup(request);
        return interactiveResponse.accessToken;
      }
      throw new Error('Failed to acquire token');
    }
  }

  private mapPosition(positionCode: number): string {
    const positions: { [key: number]: string } = {
      100000000: "Developer",
      100000001: "Manager",
      100000002: "Designer",
      100000003: "Analyst",
    };
    return positions[positionCode] || "Unknown";
  }

  private formatDate(dateStr: string | number) {
    const date = new Date(dateStr); 
    const day = String(date.getDate()).padStart(2, '0'); 
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}-${month}-${year}`; 
  }

  public async getTeamMembers(context: any): Promise<{ id: string, name: string }[]> {
    try {
      const response: HttpClientResponse = await context.httpClient.get(
        `${this.apiUrl}/new_departments/?$select=new_departmentmembers,new_departmentid`,
        HttpClient.configurations.v1,
        {
          headers: {
            Authorization: `Bearer ${await this.getAccessToken()}`,
            Accept: "application/json",
          },
        }
      );
  
      if (!response.ok) {
        throw new Error('Failed to fetch team members');
      }
  
      const data = await response.json();
      return data.value.map((item: any) => ({
        id: item.new_departmentid,   // The GUID of the team member
        name: item.new_departmentmembers, // The name of the team member
      }));
    } catch (error) {
      throw new Error('Error fetching team members: ' + error.message);
    }
  }
  

  private async fetchTeamMemberName(teamMemberId: string, context: any): Promise<string> {
    const teamMemberEndpoint = `${this.apiUrl}/new_departments(${teamMemberId})?$select=new_departmentmembers`;
    try {
      const response = await context.httpClient.get(
        teamMemberEndpoint,
        HttpClient.configurations.v1,
        {
          headers: {
            Authorization: `Bearer ${await this.getAccessToken()}`,
            Accept: "application/json",
          },
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch team member name for ID: ${teamMemberId}`);
      }

      const data = await response.json();
      return data.new_departmentmembers || "Unknown";

    } catch (error) {
      console.error(`Error fetching team member name: ${error.message}`);
      return "Unknown";
    }
  }

  public async getEmployees(context: any): Promise<IEmployee[]> {
    try {
      const accessToken = await this.getAccessToken();
      const endpointUrl = `${this.apiUrl}/new_employees`;

      const response = await context.httpClient.get(endpointUrl, HttpClient.configurations.v1, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json",
        },
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const jsonResponse = await response.json();

      const employees = jsonResponse.value.map(async (item: any) => ({
        id: item.new_employeeid,
        firstname: item.new_firstname,
        lastname: item.new_lastname,
        notes: item.new_notes,
        title: item.new_position ? this.mapPosition(item.new_position) : "",
        age: item.new_age,
        emailaddress: item.new_email,
        hiredate: item.new_hiredate ? this.formatDate(item.new_hiredate) : undefined,
        goodattheirwork: item.new_goodatthierwork || false,
        selectedTeamMember: item._new_teammember_value? await this.fetchTeamMemberName(item._new_teammember_value, context): "Unknown",
      }));

      const result = await Promise.all(employees);
      return result;
    } catch (error) {
      console.error("Error fetching employees:", error);
      throw new Error("Failed to fetch employees");
    }
  }

  public async deleteEmployee(context: any, employeeId: string): Promise<void> {
    try {
      const accessToken = await this.getAccessToken();
      const deleteEndpoint = `${this.apiUrl}/new_employees(${employeeId})`;

      const response = await context.httpClient.fetch(deleteEndpoint, HttpClient.configurations.v1, {
        method: "DELETE",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json",
        },
      });

      if (!response.ok) {
        throw new Error(`Failed to delete employee with ID: ${employeeId}`);
      }
    } catch (error) {
      console.error(`Error deleting employee: ${error.message}`);
      throw error;
    }
  }

  public async fetchEmployeeById(context: any, employeeId: string): Promise<IEmployee> {
    try {
      const accessToken = await this.getAccessToken();
      const endpointUrl = `${this.apiUrl}/new_employees(${employeeId})`;

      const response = await context.httpClient.get(endpointUrl, HttpClient.configurations.v1, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json",
        },
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const jsonResponse = await response.json();

      return {
        id: jsonResponse.new_employeeid,
        firstname: jsonResponse.new_firstname,
        lastname: jsonResponse.new_lastname,
        notes: jsonResponse.new_notes,
        title: jsonResponse.new_position ? this.mapPosition(jsonResponse.new_position) : "",
        age: jsonResponse.new_age,
        emailaddress: jsonResponse.new_email,
        hiredate: jsonResponse.new_hiredate, //yyyy-mm-ddT00:00:00Z
        goodattheirwork: jsonResponse.new_goodatthierwork || false,
        selectedTeamMember: jsonResponse._new_teammember_value ? await this.fetchTeamMemberName(jsonResponse._new_teammember_value, context) : "Unknown",
      };

    } catch (error) {
      console.error("Error fetching employee:", error);
      throw new Error("Failed to fetch employee");
    }
  }

  private isValidDate(date: string): boolean {
    const parsedDate = new Date(date);
    return !isNaN(parsedDate.getTime()) && parsedDate >= new Date('1753-01-01');
  }

  public async editEmployee(context: any, updatedEmployeeData: IEmployee): Promise<void> {
    try {
      const validHireDate = this.isValidDate(updatedEmployeeData.hiredate)
        ? updatedEmployeeData.hiredate
        : new Date('1753-01-01').toISOString();
  
      const body = {
        new_employeeid: updatedEmployeeData.id,
        new_firstname: updatedEmployeeData.firstname,
        new_lastname: updatedEmployeeData.lastname,
        new_notes: updatedEmployeeData.notes,
        new_position: updatedEmployeeData.title,
        new_age: updatedEmployeeData.age,
        new_email: updatedEmployeeData.emailaddress,
        new_hiredate: validHireDate,
        new_goodatthierwork: updatedEmployeeData.goodattheirwork,
        "new_TeamMember@odata.bind": `/new_departments(${updatedEmployeeData.selectedTeamMember})`
      };

      const accessToken = await this.getAccessToken();
      const endpointUrl = `${this.apiUrl}/new_employees(${updatedEmployeeData.id})`;
  
      const response = await context.httpClient.fetch(endpointUrl, HttpClient.configurations.v1, {
        method: "PATCH",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json",
          "Content-Type": "application/json",
        },
        body: JSON.stringify(body),
      });
  
      if (!response.ok) {
        const errorMessage = await response.text();
        throw new Error(`Failed to update employee. Response: ${errorMessage}`);
      }
  
      console.log('Employee updated successfully');
    } catch (error) {
      console.error(`Error editing employee: ${error.message}`);
      throw new Error(`Failed to edit employee: ${error.message}`);
    }
  }

  public async addEmployee(employee: Omit<IEmployee, "id">): Promise<void> {
    const token = await this.getAccessToken(); 
    const headers = {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token}`,
    };

    const validHireDate = this.isValidDate(employee.hiredate) 
    ? employee.hiredate 
    : new Date('1753-01-01').toISOString();

    const body = {
      new_firstname: employee.firstname,
      new_lastname: employee.lastname,
      new_notes: employee.notes,
      new_position: employee.title,
      new_age: employee.age,
      new_email: employee.emailaddress,
      new_hiredate: validHireDate,
      new_goodatthierwork: employee.goodattheirwork,
      "new_TeamMember@odata.bind": `/new_departments(${employee.selectedTeamMember})`
    };

    try {
      const response = await fetch(`${this.apiUrl}/new_employees`, {
        method: "POST",
        headers,
        body: JSON.stringify(body),
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(`Error adding employee: ${errorData.error.message}`);
      }
    } catch (error) {
      throw new Error(`Failed to add employee: ${error.message}`);
    }
  }

}

export default ApiService;
