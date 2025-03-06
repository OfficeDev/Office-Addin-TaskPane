import * as React from "react";
import { tokens, makeStyles } from "@fluentui/react-components";

export interface HeroListItem {
  icon: React.JSX.Element;
  primaryText: string;
}

export interface HeroListProps {
  message: string;
  items: HeroListItem[];
}

const useStyles = makeStyles({
  list: {
    marginTop: "20px",
  },
  listItem: {
    paddingBottom: "20px",
    display: "flex",
  },
  icon: {
    marginRight: "10px",
  },
  itemText: {
    fontSize: tokens.fontSizeBase300,
    fontColor: tokens.colorNeutralBackgroundStatic,
  },
  welcome__main: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  message: {
    fontSize: tokens.fontSizeBase500,
    fontColor: tokens.colorNeutralBackgroundStatic,
    fontWeight: tokens.fontWeightRegular,
    paddingLeft: "10px",
    paddingRight: "10px",
  },
});

const HeroList: React.FC<HeroListProps> = (props: HeroListProps) => {
  const { items, message } = props;
  const styles = useStyles();

  const listItems = items.map((item, index) => (
    <li className={styles.listItem} key={index}>
      <i className={styles.icon}>{item.icon}</i>
      <span className={styles.itemText}>{item.primaryText}</span>
    </li>
  ));
  return (
    <div className={styles.welcome__main}>
      <h2 className={styles.message}>{message}</h2>
      <ul className={styles.list}>{listItems}</ul>
    </div>
  );
};

export default HeroList;
