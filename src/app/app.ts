import { Component, inject, signal } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { MsalService } from '@azure/msal-angular';
import { MembersService } from './services/members-service';
import { IntegerDropdownDto } from './models/integer-dropdown-dto';
import { HttpErrorResponse } from '@angular/common/http';

@Component({
  imports: [RouterOutlet],
  selector: 'app-root',
  styleUrl: './app.scss',
  templateUrl: './app.html',
})
export class App {

  protected readonly title = signal('TiseMsalAuthTestProject');

  members: Array<IntegerDropdownDto> = [];
  defaultMember: IntegerDropdownDto = { Id: 0, Name: 'All Members' };

  constructor(private membersService: MembersService,
              private msalService: MsalService) { }

  ngOnInit(): void {
    // Required to process the redirect response from Entra ID
    this.msalService.handleRedirectObservable().subscribe({
      next: (result) => {
        if (result?.account) {
          this.msalService.instance.setActiveAccount(result.account);
        }
      },
      error: (error) => console.error('MSAL Redirect Error:', error),
    });
  }

  login(): void {
    this.msalService.loginRedirect();
  }

  logout(): void {
    this.msalService.logoutRedirect();
  }



  getMembers(): void {
        this.membersService.getMembers().subscribe(
          { 
              next: (result: IntegerDropdownDto[]) => {   
                this.members = [this.defaultMember, ...result];         
                console.log('members: ', this.members);
              },
            
              error: (error: HttpErrorResponse) => {
              console.error('Error loading members:', error);
            }
          }
        );
    }
}
