import * as React from 'react';
import { useEffect, useState } from 'react';
import Pagination from 'react-js-pagination';
import styles from './Paging.module.scss';
import { Icon } from '@fluentui/react';

export type PageUpdateCallback = (pageNumber: number) => void;

export interface IPagingProps {
	totalItems: number;
	itemsCountPerPage: number;
	onPageUpdate: PageUpdateCallback;
	currentPage: number;
	pageRange?: number;
}

export interface IPagingState {
	currentPage: number;
}

const Paging: React.FC<IPagingProps> = (props) => {
	const [currentPage, setcurrentPage] = useState<number>(props.currentPage);

	const _pageChange = (pageNumber: number): void => {
		setcurrentPage(pageNumber);
		props.onPageUpdate(pageNumber);
	};

	useEffect(() => {
		setcurrentPage(props.currentPage);
	}, [props.currentPage]);

	return (
		<div className={styles.paginationContainer}>
			<div className={styles.searchWp__paginationContainer__pagination}>
				<Pagination
					activePage={currentPage}
					firstPageText={currentPage === 1 ? null :
						<Icon iconName="ChevronLeftEnd6" />}
					lastPageText={<Icon iconName="ChevronRightEnd6" />}
					prevPageText={currentPage === 1 ? null : <Icon iconName="ChevronLeft" />}
					nextPageText={<Icon iconName="ChevronRight" />}
					activeLinkClass={styles.active}
					itemsCountPerPage={props.itemsCountPerPage}
					totalItemsCount={props.totalItems}
					pageRangeDisplayed={props.pageRange || 3}
					onChange={_pageChange}
				/>
			</div>
		</div>
	);
};

export default Paging;
