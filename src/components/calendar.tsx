import * as React from "react";
import { Event } from "microsoft-graph";
import moment from "moment";

function formatDateTime(dateTime: string | undefined) {
  if (dateTime !== undefined) {
    return moment.utc(dateTime).local().format("M/D/YY h:mm A");
  }
}

class Calendar extends React.Component<{ events: any }> {
  render() {
    return (
      <div>
        <h1>Calendar</h1>
        <table>
          <thead>
            <tr>
              <th scope="col">Organizer</th>
              <th scope="col">Subject</th>
              <th scope="col">Start</th>
              <th scope="col">End</th>
            </tr>
          </thead>
          <tbody>
            {this.props.events.map(function (event: Event) {
              return (
                <tr key={event.id}>
                  <td>{event.organizer?.emailAddress?.name}</td>
                  <td>{event.subject}</td>
                  <td>{formatDateTime(event.start?.dateTime)}</td>
                  <td>{formatDateTime(event.end?.dateTime)}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  }
}
export default Calendar;
